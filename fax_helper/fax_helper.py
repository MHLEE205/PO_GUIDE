# -*- coding: utf-8 -*-
"""
POFaxHelper — P/O申請ガイド用 FAX自動入力ヘルパー

PO_GUIDE（https://mhlee205.github.io/PO_GUIDE/）で作成したPDFを受け取り、
複合機のFAXプリンター（例: ApeosPort ... FAX）へ印刷指示 → 「ファクス送信の設定／確認」画面に
FAX番号・宛先名を自動入力する。「送信開始」はユーザーが画面で確認してから押す運用（自動では押さない）。

仕組みは sample/FAX送信実装マニュアル.docx（booking_mailer.py）の
_find_fax_dialog / _autofill_fax_one / _fax_send_sequential を移植したもの。

- 127.0.0.1:18765 でHTTP待ち受け（PO_GUIDEのOriginからのリクエストのみ受け付ける）
- 初回起動時にWindowsのスタートアップ（HKCU\\...\\Run）へ自動登録
- タスクトレイ常駐（終了・スタートアップ解除はトレイメニューから）
"""
import base64
import json
import logging
import os
import re
import sys
import tempfile
import threading
import time
import urllib.request
import uuid
import winreg
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import win32api
import win32clipboard
import win32con
import win32gui
import win32print

VERSION = '1.0.0'
PORT = 18765
ALLOWED_ORIGINS = {'https://mhlee205.github.io'}
PROFILES_URL = 'https://mhlee205.github.io/PO_GUIDE/fax_helper/profiles.json'
APP_NAME = 'POFaxHelper'
APP_DIR = os.path.join(os.environ.get('APPDATA', tempfile.gettempdir()), APP_NAME)
CONFIG_PATH = os.path.join(APP_DIR, 'config.json')
os.makedirs(APP_DIR, exist_ok=True)

logging.basicConfig(
    filename=os.path.join(APP_DIR, 'helper.log'), level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s', encoding='utf-8')
log = logging.getLogger(APP_NAME)

# FAX画面ごとの入力手順。サーバー（profiles.json）から取得できない場合はこの内蔵定義を使う。
# 新しい形のFAX画面が見つかったら profiles.json に追加すれば、exeを再配布せずに対応できる。
BUILTIN_PROFILES = {
    'profiles': [
        {
            'name': 'FUJIFILM ApeosPort FAXドライバー',
            'title_contains': ['ファクス送信の設定'],
            'steps': [
                ['key', 'alt+f'], ['key', 'ctrl+a'], ['key', 'delete'], ['type', '{number}'],
                ['key', 'alt+n'], ['key', 'ctrl+a'], ['key', 'delete'], ['paste', '{name}'],
                ['key', 'alt+i'],
            ],
        }
    ],
    # どのプロファイルにも一致しないが、FAX画面らしいタイトル（未対応の画面）
    'generic_title_contains': ['ファクス', 'FAX', 'Fax'],
}

_profiles_cache = {'data': None, 'at': 0}


def load_profiles():
    if _profiles_cache['data'] and time.time() - _profiles_cache['at'] < 600:
        return _profiles_cache['data']
    data = BUILTIN_PROFILES
    try:
        with urllib.request.urlopen(PROFILES_URL, timeout=5) as r:
            remote = json.loads(r.read().decode('utf-8'))
        if remote.get('profiles'):
            data = remote
    except Exception as e:
        log.warning('profiles.json取得失敗（内蔵定義を使用）: %s', e)
    _profiles_cache.update(data=data, at=time.time())
    return data


# ── 設定（選択したFAXプリンター） ──
def load_config():
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def list_fax_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    names = [p[2] for p in win32print.EnumPrinters(flags)]
    # Windows標準の「Fax」（FAXとスキャン）はモデム/FAXサーバーが無いと送れないため除外
    return [n for n in names if 'FAX' in n.upper() and n.strip().upper() != 'FAX']


def current_printer():
    printers = list_fax_printers()
    saved = load_config().get('printer')
    if saved in printers:
        return saved, printers
    # 複合機（ApeosPort / FUJIFILM）のFAXを優先
    preferred = next((p for p in printers if 'APEOS' in p.upper() or p.upper().startswith('FF ')), None)
    return (preferred or (printers[0] if printers else None)), printers


# ── キーボード操作 ──
VK = {'alt': win32con.VK_MENU, 'ctrl': win32con.VK_CONTROL, 'shift': win32con.VK_SHIFT,
      'delete': win32con.VK_DELETE, 'enter': win32con.VK_RETURN, 'tab': win32con.VK_TAB}


def _vk(name):
    name = name.lower()
    if name in VK:
        return VK[name]
    if len(name) == 1:
        return ord(name.upper())
    raise ValueError(f'不明なキー: {name}')


def press(combo):
    keys = [_vk(k) for k in combo.split('+')]
    for k in keys:
        win32api.keybd_event(k, 0, 0, 0)
    for k in reversed(keys):
        win32api.keybd_event(k, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.15)


def type_digits(text):
    for ch in text:
        if ch.isdigit():
            win32api.keybd_event(ord(ch), 0, 0, 0)
            win32api.keybd_event(ord(ch), 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.05)
    time.sleep(0.1)


def get_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass
    return None


def set_clipboard_text(text):
    for _ in range(10):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()
            return True
        except Exception:
            time.sleep(0.1)
    return False


def paste_text(text):
    # 日本語はキー入力できないためクリップボード経由で貼り付け（元のクリップボード内容は復元）
    prev = get_clipboard_text()
    set_clipboard_text(text)
    time.sleep(0.1)
    press('ctrl+v')
    time.sleep(0.2)
    if prev is not None:
        set_clipboard_text(prev)


def focus_window(hwnd):
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # 別プロセスのウィンドウを前面にするための定番の回避策（Altキーを一度押す）
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)
    except Exception as e:
        log.warning('前面化失敗: %s', e)
    return win32gui.GetForegroundWindow() == hwnd


# ── FAX画面の検出 ──
def find_fax_dialog(profiles):
    """(hwnd, profile or None, title) を返す。profile=None は未対応のFAX画面"""
    found = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return True
        for p in profiles['profiles']:
            if any(k in title for k in p['title_contains']):
                found.append((0, hwnd, p, title))
                return True
        if any(k in title for k in profiles.get('generic_title_contains', [])) and 'PO_GUIDE' not in title \
                and 'P/O申請ガイド' not in title:
            found.append((1, hwnd, None, title))
        return True

    win32gui.EnumWindows(cb, None)
    if not found:
        return None, None, None
    found.sort(key=lambda x: x[0])
    _, hwnd, p, title = found[0]
    return hwnd, p, title


def window_exists(hwnd):
    return win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd)


# ── ジョブ ──
JOBS = {}
JOBS_LOCK = threading.Lock()


def set_job(job_id, **kw):
    with JOBS_LOCK:
        JOBS[job_id].update(kw)
    log.info('job %s: %s', job_id, kw)


def run_fax_job(job_id, pdf_path, numbers, name, printer):
    try:
        profiles = load_profiles()
        # 既に開いているFAX画面を誤って操作しないよう、印刷前に存在していたものを記録
        before, _, _ = find_fax_dialog(profiles)

        set_job(job_id, state='printing', message='FAXプリンターへ送っています')
        win32api.ShellExecute(0, 'printto', pdf_path, f'"{printer}"', os.path.dirname(pdf_path), 1)
        time.sleep(2)

        hwnd = profile = title = None
        for _ in range(60):  # 最大30秒
            h, p, t = find_fax_dialog(profiles)
            if h and h != before:
                hwnd, profile, title = h, p, t
                break
            time.sleep(0.5)
        if not hwnd:
            set_clipboard_text(numbers[0])
            set_job(job_id, state='manual',
                    message='FAX画面が見つかりませんでした。FAX番号をクリップボードにコピーしたので、FAX画面が開いたら番号欄に貼り付けてください。')
            return

        time.sleep(0.8)  # 画面の描画完了待ち
        if profile is None:
            set_clipboard_text(numbers[0])
            set_job(job_id, state='manual',
                    message=f'未対応のFAX画面です（{title}）。FAX番号をクリップボードにコピーしたので、番号欄に貼り付けて送信してください。')
        else:
            for number in numbers:
                if not window_exists(hwnd):
                    raise RuntimeError('入力中にFAX画面が閉じられました')
                if not focus_window(hwnd):
                    # 前面化に失敗したまま入力すると別ウィンドウに打鍵してしまうため中止
                    set_clipboard_text(number)
                    set_job(job_id, state='manual',
                            message='FAX画面を前面に表示できませんでした。FAX番号をクリップボードにコピーしたので、番号欄に貼り付けてください。')
                    break
                for step in profile['steps']:
                    kind, arg = step[0], step[1].replace('{number}', number).replace('{name}', name)
                    if kind == 'key':
                        press(arg)
                    elif kind == 'type':
                        type_digits(arg)
                    elif kind == 'paste':
                        paste_text(arg)
                    elif kind == 'wait':
                        time.sleep(float(arg))
            else:
                set_job(job_id, state='filled',
                        message='FAX番号・宛先名を入力しました。FAX画面で内容を確認し「送信開始」を押してください。')

        for _ in range(1200):  # 最大10分、ユーザーが送信開始/中止するまで待つ
            if not window_exists(hwnd):
                set_job(job_id, state='closed', message='FAX画面が閉じられました（送信開始 または 送信中止）')
                return
            time.sleep(0.5)
        set_job(job_id, state='timeout', message='10分経過したため監視を終了しました')
    except Exception as e:
        log.exception('job %s failed', job_id)
        set_job(job_id, state='error', message=str(e))
    finally:
        # Acrobat等がファイルを掴んでいる可能性があるため、しばらく後に削除
        threading.Timer(600, lambda: _safe_remove(pdf_path)).start()


def _safe_remove(path):
    try:
        os.remove(path)
    except Exception:
        pass


# ── HTTPサーバー ──
class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        log.info('%s - %s', self.address_string(), fmt % args)

    def _origin_ok(self):
        return self.headers.get('Origin') in ALLOWED_ORIGINS

    def _send(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        if self._origin_ok():
            self.send_header('Access-Control-Allow-Origin', self.headers['Origin'])
            self.send_header('Vary', 'Origin')
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204 if self._origin_ok() else 403)
        if self._origin_ok():
            self.send_header('Access-Control-Allow-Origin', self.headers['Origin'])
            self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type')
            # Chrome/EdgeのPrivate Network Access（公開サイト→ローカル）対応
            self.send_header('Access-Control-Allow-Private-Network', 'true')
            self.send_header('Vary', 'Origin')
        self.end_headers()

    def _read_json(self):
        n = int(self.headers.get('Content-Length') or 0)
        if n > 30 * 1024 * 1024:
            raise ValueError('データが大きすぎます')
        return json.loads(self.rfile.read(n).decode('utf-8') or '{}')

    def do_GET(self):
        if not self._origin_ok():
            return self._send(403, {'error': 'forbidden'})
        if self.path == '/status':
            printer, printers = current_printer()
            return self._send(200, {'ok': True, 'version': VERSION, 'printer': printer, 'printers': printers})
        m = re.fullmatch(r'/job/([0-9a-f]+)', self.path)
        if m:
            with JOBS_LOCK:
                job = dict(JOBS.get(m.group(1)) or {})
            return self._send(200 if job else 404, job or {'error': 'not found'})
        self._send(404, {'error': 'not found'})

    def do_POST(self):
        if not self._origin_ok():
            return self._send(403, {'error': 'forbidden'})
        try:
            data = self._read_json()
        except Exception as e:
            return self._send(400, {'error': str(e)})

        if self.path == '/config':
            printer = data.get('printer')
            if printer not in list_fax_printers():
                return self._send(400, {'error': 'プリンターが見つかりません'})
            cfg = load_config()
            cfg['printer'] = printer
            save_config(cfg)
            return self._send(200, {'ok': True, 'printer': printer})

        if self.path == '/fax':
            numbers = [re.sub(r'\D', '', str(n)) for n in (data.get('fax_numbers') or [])]
            numbers = [n for n in numbers if len(n) >= 9]
            if not numbers:
                return self._send(400, {'error': 'FAX番号がありません'})
            printer, _ = current_printer()
            if not printer:
                return self._send(400, {'error': 'FAXプリンターが見つかりません（名前に「FAX」を含むプリンターを登録してください）'})
            try:
                pdf = base64.b64decode(data.get('pdf_base64') or '')
            except Exception:
                return self._send(400, {'error': 'PDFデータが不正です'})
            if not pdf.startswith(b'%PDF-'):
                return self._send(400, {'error': 'PDFデータが不正です'})
            with JOBS_LOCK:
                busy = any(j['state'] in ('printing', 'filled') for j in JOBS.values())
            if busy:
                return self._send(409, {'error': '前のFAX画面がまだ開いています。送信開始または送信中止してから再度お試しください。'})

            safe_name = re.sub(r'[\\/:*?"<>|]', '_', data.get('filename') or 'PO.pdf')
            job_id = uuid.uuid4().hex
            job_dir = os.path.join(tempfile.gettempdir(), APP_NAME, job_id)
            os.makedirs(job_dir, exist_ok=True)
            pdf_path = os.path.join(job_dir, safe_name)
            with open(pdf_path, 'wb') as f:
                f.write(pdf)
            name = str(data.get('recipient') or '')[:60]
            with JOBS_LOCK:
                JOBS[job_id] = {'state': 'queued', 'message': '', 'printer': printer, 'numbers': numbers}
            threading.Thread(target=run_fax_job, args=(job_id, pdf_path, numbers, name, printer), daemon=True).start()
            return self._send(200, {'ok': True, 'job_id': job_id, 'printer': printer})

        self._send(404, {'error': 'not found'})


# ── スタートアップ登録 ──
RUN_KEY = r'Software\Microsoft\Windows\CurrentVersion\Run'


def exe_path():
    return sys.executable if getattr(sys, 'frozen', False) else None


def is_startup_registered():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP_NAME)
            return True
    except OSError:
        return False


def register_startup():
    path = exe_path()
    if not path:
        return
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, f'"{path}"')


def unregister_startup():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, APP_NAME)
    except OSError:
        pass


def main():
    try:
        server = ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    except OSError:
        # 既に起動済み（ポート使用中）なら何もせず終了
        log.info('already running')
        return
    cfg = load_config()
    if not cfg.get('startup_asked'):
        register_startup()
        cfg['startup_asked'] = True
        save_config(cfg)
    log.info('POFaxHelper %s started on %d', VERSION, PORT)
    threading.Thread(target=server.serve_forever, daemon=True).start()

    try:
        import pystray
        from PIL import Image, ImageDraw
    except ImportError:
        # 開発時（トレイ用ライブラリ無し）はそのまま待ち受け続ける
        print(f'POFaxHelper {VERSION} listening on 127.0.0.1:{PORT} (Ctrl+C to quit)')
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            return

    img = Image.new('RGB', (64, 64), (43, 91, 215))
    d = ImageDraw.Draw(img)
    d.rectangle([12, 22, 52, 50], fill=(255, 255, 255))
    d.rectangle([20, 12, 44, 24], fill=(255, 255, 255))
    d.rectangle([18, 30, 46, 34], fill=(43, 91, 215))

    def toggle_startup(icon, item):
        unregister_startup() if is_startup_registered() else register_startup()

    def quit_app(icon, item):
        icon.stop()
        server.shutdown()

    menu = pystray.Menu(
        pystray.MenuItem(f'POFaxHelper v{VERSION}', None, enabled=False),
        pystray.MenuItem(lambda item: f'FAXプリンター: {current_printer()[0] or "なし"}', None, enabled=False),
        pystray.MenuItem('Windows起動時に自動起動', toggle_startup, checked=lambda item: is_startup_registered()),
        pystray.MenuItem('終了', quit_app),
    )
    pystray.Icon(APP_NAME, img, 'POFaxHelper（P/O FAX自動入力）', menu).run()


if __name__ == '__main__':
    main()
