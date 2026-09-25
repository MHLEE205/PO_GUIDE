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
import win32event
import win32gui
import win32print

VERSION = '1.2.0'
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


def run_fax_job(job_id, pdf_path, numbers, name, printer, delete_after=True):
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
        # 一時ファイルのみ削除（Acrobat等がファイルを掴んでいる可能性があるため、しばらく後に）
        if delete_after:
            threading.Timer(600, lambda: _safe_remove(pdf_path)).start()


def _safe_remove(path):
    try:
        os.remove(path)
    except Exception:
        pass


# ── ファイル関連 ──
DOCS = {}  # doc_id → {path, numbers, recipient, job_id, dialog}（/open で保存・表示したPDF）


def decode_pdf(data):
    try:
        pdf = base64.b64decode(data.get('pdf_base64') or '')
    except Exception:
        return None, 'PDFデータが不正です'
    if not pdf.startswith(b'%PDF-'):
        return None, 'PDFデータが不正です'
    return pdf, None


def safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|]', '_', str(name or 'PO.pdf')).strip() or 'PO.pdf'
    return name if name.lower().endswith('.pdf') else name + '.pdf'


def downloads_dir():
    # ダウンロードフォルダの実際の場所（OneDrive等へ移動されている場合も含む）をWindowsから取得
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as k:
            path = os.path.expandvars(winreg.QueryValueEx(k, '{374DE290-123F-4565-9164-39C4925E467B}')[0])
            if os.path.isdir(path):
                return path
    except OSError:
        pass
    path = os.path.join(os.path.expanduser('~'), 'Downloads')
    os.makedirs(path, exist_ok=True)
    return path


def unique_path(folder, filename):
    # ブラウザのダウンロードと同じく、同名ファイルがあれば「名前 (1).pdf」形式で連番を付ける
    base, ext = os.path.splitext(filename)
    path = os.path.join(folder, filename)
    n = 1
    while os.path.exists(path):
        path = os.path.join(folder, f'{base} ({n}){ext}')
        n += 1
    return path


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
        # /open で開いたPDFの状態（確認ダイアログから開始したFAXのjob_idを含む。サイト側がポーリングして進捗表示に使う）
        m = re.fullmatch(r'/doc/([0-9a-f]+)', self.path)
        if m:
            with JOBS_LOCK:
                doc = DOCS.get(m.group(1))
                info = {'job_id': doc.get('job_id'), 'dialog': doc.get('dialog')} if doc else None
            return self._send(200 if info else 404, info or {'error': 'not found'})
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

        if self.path == '/open':
            # マニュアル4-2と同じく、PDFを既定のアプリ（Acrobat/Foxit等）で開いて利用者に確認・修正してもらい、
            # 同時に最前面の確認ダイアログを表示。「確認完了 → FAX画面を開く」で保存済みの最終版をFAXする。
            # 保存先はダウンロードフォルダ（ブラウザでダウンロードしていた従来と同じ場所）
            pdf, err = decode_pdf(data)
            if err:
                return self._send(400, {'error': err})
            pdf_path = unique_path(downloads_dir(), safe_filename(data.get('filename')))
            with open(pdf_path, 'wb') as f:
                f.write(pdf)
            doc_id = uuid.uuid4().hex
            doc = {'path': pdf_path, 'job_id': None, 'dialog': None,
                   'numbers': clean_numbers(data.get('fax_numbers')),
                   'recipient': str(data.get('recipient') or '')[:60]}
            with JOBS_LOCK:
                # 前回の確認ダイアログが残っていれば閉じる（新しいP/Oに切り替わったため）
                for old in DOCS.values():
                    if old.get('dialog') == 'open':
                        old['dialog'] = 'close_requested'
                DOCS[doc_id] = doc
            opened = True
            try:
                os.startfile(pdf_path)
            except Exception as e:
                log.warning('PDFを開けませんでした: %s', e)
                opened = False
            if doc['numbers']:
                doc['dialog'] = 'open'
                threading.Thread(target=show_confirm_dialog, args=(doc_id,), daemon=True).start()
            return self._send(200, {'ok': True, 'doc_id': doc_id, 'path': pdf_path, 'opened': opened})

        if self.path == '/fax':
            numbers = clean_numbers(data.get('fax_numbers'))
            name = str(data.get('recipient') or '')[:60]
            if data.get('doc_id'):
                # /open で開いたファイル（利用者が修正・保存済みの場合はその最終版）をそのまま送る
                job_id, err, code = start_doc_fax(data['doc_id'], numbers, name)
            else:
                pdf, err = decode_pdf(data)
                if err:
                    return self._send(400, {'error': err})
                tmp_dir = os.path.join(tempfile.gettempdir(), APP_NAME, uuid.uuid4().hex)
                os.makedirs(tmp_dir, exist_ok=True)
                pdf_path = os.path.join(tmp_dir, safe_filename(data.get('filename')))
                with open(pdf_path, 'wb') as f:
                    f.write(pdf)
                job_id, err, code = start_fax_job(pdf_path, numbers, name, delete_after=True)
            if err:
                return self._send(code, {'error': err})
            return self._send(200, {'ok': True, 'job_id': job_id})

        self._send(404, {'error': 'not found'})


def clean_numbers(values):
    numbers = [re.sub(r'\D', '', str(n)) for n in (values or [])]
    return [n for n in numbers if len(n) >= 9]


def start_fax_job(pdf_path, numbers, name, delete_after):
    """FAXジョブを開始して (job_id, エラー文, HTTPコード) を返す"""
    if not numbers:
        return None, 'FAX番号がありません', 400
    printer, _ = current_printer()
    if not printer:
        return None, 'FAXプリンターが見つかりません（名前に「FAX」を含むプリンターを登録してください）', 400
    with JOBS_LOCK:
        if any(j['state'] in ('queued', 'printing', 'filled') for j in JOBS.values()):
            return None, '前のFAX画面がまだ開いています。送信開始または送信中止してから再度お試しください。', 409
        job_id = uuid.uuid4().hex
        JOBS[job_id] = {'state': 'queued', 'message': '', 'printer': printer, 'numbers': numbers}
    threading.Thread(target=run_fax_job, args=(job_id, pdf_path, numbers, name, printer, delete_after), daemon=True).start()
    return job_id, None, 200


def start_doc_fax(doc_id, numbers=None, name=None):
    """/open で開いたPDFをFAXする（サイトのボタン・確認ダイアログ共通）。利用者のファイルなので削除しない"""
    with JOBS_LOCK:
        doc = DOCS.get(doc_id)
    if not doc or not os.path.exists(doc['path']):
        return None, 'PDFファイルが見つかりません（移動・削除された可能性があります）。P/O自動作成からやり直してください。', 400
    with open(doc['path'], 'rb') as f:
        if not f.read(5).startswith(b'%PDF-'):
            return None, 'PDFファイルが不正です', 400
    numbers = numbers or doc['numbers']
    name = name if name is not None else doc['recipient']
    job_id, err, code = start_fax_job(doc['path'], numbers, name, delete_after=False)
    if job_id:
        with JOBS_LOCK:
            doc['job_id'] = job_id
            if doc.get('dialog') == 'open':
                doc['dialog'] = 'close_requested'  # サイトのボタンで開始した場合は確認ダイアログを閉じる
    return job_id, err, code


# ── 確認ダイアログ（マニュアル4-2のPDF確認ダイアログに相当） ──
def show_confirm_dialog(doc_id):
    import tkinter as tk

    with JOBS_LOCK:
        doc = DOCS[doc_id]
    font = ('Yu Gothic UI', 10)
    root = tk.Tk()
    root.title('POFaxHelper - FAX送信の確認')
    root.attributes('-topmost', True)
    root.resizable(False, False)
    root.configure(padx=14, pady=12)

    tk.Label(root, text=f'📄 {os.path.basename(doc["path"])} を開きました', font=('Yu Gothic UI', 10, 'bold'),
             anchor='w', justify='left').pack(fill='x')
    tk.Label(root, text=f'送信先: {" / ".join(doc["numbers"])}　{doc["recipient"]}', font=font,
             anchor='w', justify='left', wraplength=420).pack(fill='x', pady=(4, 0))
    tk.Label(root, text='PDFを修正した場合は、先にPDFの画面で保存（Ctrl+S）してから押してください。',
             font=('Yu Gothic UI', 9), fg='#b45309', anchor='w', justify='left', wraplength=420).pack(fill='x', pady=(6, 8))
    msg = tk.Label(root, text='', font=('Yu Gothic UI', 9), fg='#b91c1c', anchor='w', justify='left', wraplength=420)
    btns = tk.Frame(root)
    btns.pack(fill='x')

    def close(state='closed'):
        with JOBS_LOCK:
            doc['dialog'] = state
        root.destroy()

    def on_fax():
        # FAX画面への自動キー入力がこのダイアログに入らないよう、先に閉じてから開始する
        root.withdraw()
        job_id, err, _ = start_doc_fax(doc_id)
        if err:
            root.deiconify()
            msg.config(text=err)
            msg.pack(fill='x', pady=(8, 0))
            return
        close('done')

    tk.Button(btns, text='📠 確認完了 → FAX画面を開く', font=('Yu Gothic UI', 10, 'bold'), bg='#16a34a', fg='white',
              activebackground='#15803d', activeforeground='white', padx=12, pady=4, command=on_fax).pack(side='left')
    tk.Button(btns, text='取消', font=font, padx=10, pady=4, command=close).pack(side='left', padx=(8, 0))
    root.protocol('WM_DELETE_WINDOW', close)

    # 画面中央に表示（見落とさないように）
    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    root.geometry(f'+{(root.winfo_screenwidth() - w) // 2}+{(root.winfo_screenheight() - h) // 2}')

    def watch():
        # サイト側のボタンでFAXを開始した・新しいP/Oが開かれた場合は自動で閉じる
        with JOBS_LOCK:
            requested = doc.get('dialog') == 'close_requested'
        if requested:
            close('closed')
        else:
            root.after(300, watch)

    root.after(300, watch)
    root.mainloop()


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


class ExclusiveServer(ThreadingHTTPServer):
    # Windowsでは SO_REUSEADDR を付けると同じポートを複数プロセスが同時にbindできてしまい、
    # 多重起動を検出できないため無効化する
    allow_reuse_address = False


def notify_already_running():
    win32api.MessageBox(0, 'POFaxHelperは既に起動しています。\n画面右下のタスクトレイ（^）に常駐しています。',
                        'POFaxHelper', win32con.MB_OK | win32con.MB_ICONINFORMATION)


def main():
    # 多重起動防止（exeを何度もダブルクリックしても1つだけ動くようにする）
    mutex = win32event.CreateMutex(None, False, 'Local\\POFaxHelper_SingleInstance')
    if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        log.info('already running (mutex)')
        notify_already_running()
        return
    try:
        server = ExclusiveServer(('127.0.0.1', PORT), Handler)
    except OSError:
        log.info('already running (port in use)')
        notify_already_running()
        return
    cfg = load_config()
    # 初回は自動登録。登録済みの場合も、exeを移動・更新したときに備えて現在のexeパスで登録し直す
    if not cfg.get('startup_asked') or is_startup_registered():
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
    def on_ready(icon):
        icon.visible = True
        # 画面を持たない常駐ツールのため、起動したことが分かるよう通知を出す
        try:
            icon.notify('起動しました。タスクトレイに常駐し、PO_GUIDEからのFAX自動入力を待ち受けます。', 'POFaxHelper')
        except Exception:
            pass

    pystray.Icon(APP_NAME, img, 'POFaxHelper（P/O FAX自動入力）', menu).run(setup=on_ready)
    del mutex


if __name__ == '__main__':
    main()
