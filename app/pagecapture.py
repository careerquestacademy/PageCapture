"""
PageCapture v3.1
Desktop app — receives page content from Firefox extension via Native Messaging,
converts to clean editable DOCX with embedded images and clickable links.
No keyboard hooks. No mouse hooks. No system takeover. Ever.
"""

import sys
import os
import json
import struct
import threading
import tkinter as tk
from tkinter import filedialog
import re
import base64
import io

# ── Graceful imports ──────────────────────────────────────────────────────────

try:
    from docx import Document
    from docx.shared import RGBColor, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

import sys as _sys
HAS_WIN32 = _sys.platform == "win32"

# ── Config ────────────────────────────────────────────────────────────────────

APP_NAME    = "PageCapture"
APP_VERSION = "3.1"
SAVE_FOLDER = os.path.expanduser("~/Documents/PageCapture")
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".pagecapture3_config.json")

os.makedirs(SAVE_FOLDER, exist_ok=True)

COLORS = {
    'bg':     '#1e1e2e',
    'panel':  '#313244',
    'border': '#45475a',
    'text':   '#cdd6f4',
    'muted':  '#a6adc8',
    'purple': '#cba6f7',
    'blue':   '#89b4fa',
    'green':  '#a6e3a1',
    'red':    '#f38ba8',
    'yellow': '#f9e2af',
}

def load_config():
    try:
        with open(CONFIG_PATH) as f:
            return json.load(f)
    except Exception:
        return {"save_folder": SAVE_FOLDER}

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

# ── Native Messaging I/O ──────────────────────────────────────────────────────
# Firefox communicates via stdin/stdout using 4-byte little-endian length prefix.
# Because Firefox launches a SEPARATE instance of the exe as a subprocess,
# we use a local socket to forward captured data to the visible GUI instance.

SOCKET_PORT = 27183  # internal forwarding port (localhost only, not Firefox)

def read_native_message():
    """Read one message from Firefox via stdin. Returns None on EOF."""
    raw_len = sys.stdin.buffer.read(4)
    if len(raw_len) < 4:
        return None
    msg_len = struct.unpack("<I", raw_len)[0]
    raw_msg = sys.stdin.buffer.read(msg_len)
    return json.loads(raw_msg.decode("utf-8"))

def send_native_message(data):
    """Send one message back to Firefox via stdout."""
    encoded = json.dumps(data).encode("utf-8")
    sys.stdout.buffer.write(struct.pack("<I", len(encoded)))
    sys.stdout.buffer.write(encoded)
    sys.stdout.buffer.flush()

def is_gui_instance():
    """True if we were launched by the user (no stdin data from Firefox)."""
    return sys.stdin is None or not hasattr(sys.stdin, 'buffer')

def start_native_listener(app):
    """
    Two modes:
    1. GUI instance (user double-clicked): listen on local socket for forwarded data.
    2. Host instance (Firefox launched us): read stdin, forward to GUI via socket, exit.
    """
    import socket as _socket

    def _forward_to_gui(msg):
        """Called in host instance — send data to GUI instance via socket."""
        try:
            s = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
            s.settimeout(3)
            s.connect(('127.0.0.1', SOCKET_PORT))
            data = json.dumps(msg).encode('utf-8')
            s.sendall(struct.pack('<I', len(data)) + data)
            s.close()
            send_native_message({"status": "ok"})
        except Exception as e:
            send_native_message({"status": "error", "message": str(e)})

    def _host_mode():
        """Read from Firefox stdin and forward to GUI instance."""
        while True:
            msg = read_native_message()
            if msg is None:
                break
            _forward_to_gui(msg)

    def _gui_socket_server():
        """Listen for forwarded messages from host instances."""
        import socket as _socket
        srv = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
        srv.setsockopt(_socket.SOL_SOCKET, _socket.SO_REUSEADDR, 1)
        try:
            srv.bind(('127.0.0.1', SOCKET_PORT))
        except OSError:
            return  # another GUI instance already listening
        srv.listen(5)
        while True:
            try:
                conn, _ = srv.accept()
                raw_len = conn.recv(4)
                if len(raw_len) < 4:
                    conn.close()
                    continue
                msg_len = struct.unpack('<I', raw_len)[0]
                raw_msg = b''
                while len(raw_msg) < msg_len:
                    chunk = conn.recv(msg_len - len(raw_msg))
                    if not chunk:
                        break
                    raw_msg += chunk
                conn.close()
                msg = json.loads(raw_msg.decode('utf-8'))
                app.root.after(0, lambda m=msg: app.on_content_received(m))
            except Exception:
                continue

    # Detect whether Firefox launched us (stdin has data) or user launched us
    test = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
    test.settimeout(0.1)
    try:
        test.connect(('127.0.0.1', SOCKET_PORT))
        test.close()
        # GUI instance already running — we are the host instance
        t = threading.Thread(target=_host_mode, daemon=True)
        t.start()
    except OSError:
        test.close()
        # No GUI instance yet — we are the GUI instance, start socket server
        t = threading.Thread(target=_gui_socket_server, daemon=True)
        t.start()

# ── DOCX builder ──────────────────────────────────────────────────────────────

def blocks_to_docx(title, blocks, save_path):
    """
    Convert content blocks to a clean, fully editable DOCX.
    Text is real text. Links are clickable. Images are embedded.
    """
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    doc.add_heading(title or 'Captured Page', level=1)

    for block in blocks:
        btype    = block.get('type', 'text')
        text     = block.get('text', '').strip()
        href     = block.get('href', '')
        alt      = block.get('alt', '')
        level    = block.get('level', 2)
        src      = block.get('src', '')
        img_data = block.get('img_data', '')

        if btype == 'heading' and text:
            doc.add_heading(text, level=min(level, 6))

        elif btype == 'link' and text:
            p = doc.add_paragraph()
            _add_hyperlink(p, text, href)

        elif btype == 'listitem' and text:
            doc.add_paragraph(text, style='List Bullet')

        elif btype == 'image':
            embedded = False
            if img_data:
                try:
                    if ',' in img_data:
                        img_data = img_data.split(',', 1)[1]
                    img_bytes  = base64.b64decode(img_data)
                    img_stream = io.BytesIO(img_bytes)
                    if HAS_PIL:
                        pil_img = Image.open(img_stream)
                        w, h    = pil_img.size
                        img_stream.seek(0)
                        max_w = Inches(5)
                        if w > 0:
                            ratio = min(1.0, max_w / (w * 9144))
                            doc_w = Inches(w * ratio / 96)
                        else:
                            doc_w = Inches(4)
                        doc.add_picture(img_stream, width=doc_w)
                    else:
                        doc.add_picture(img_stream, width=Inches(4))
                    embedded = True
                except Exception:
                    embedded = False

            if not embedded:
                p   = doc.add_paragraph()
                run = p.add_run(f'[Image: {alt or src or "image"}]')
                run.italic = True
                run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

        elif btype == 'text' and text:
            doc.add_paragraph(text)

    doc.save(save_path)


def _add_hyperlink(paragraph, text, url):
    try:
        part  = paragraph.part
        r_id  = part.relate_to(
            url,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        new_run = OxmlElement('w:r')
        rPr     = OxmlElement('w:rPr')
        rStyle  = OxmlElement('w:rStyle')
        rStyle.set(qn('w:val'), 'Hyperlink')
        rPr.append(rStyle)
        new_run.append(rPr)
        new_run.text = text
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
    except Exception:
        run = paragraph.add_run(f'{text} [{url}]')
        run.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
        run.underline = True


# ── Find installed document editors ──────────────────────────────────────────

def find_doc_editors():
    editors = []
    candidates = [
        ("LibreOffice Writer",
         [r"C:\Program Files\LibreOffice\program\swriter.exe",
          r"C:\Program Files (x86)\LibreOffice\program\swriter.exe"]),
        ("Microsoft Word",
         [r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
          r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
          r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"]),
        ("Polaris Office",
         [r"C:\Program Files\Polaris Office\PolarisOffice.exe",
          r"C:\Program Files (x86)\Polaris Office\PolarisOffice.exe"]),
        ("WPS Writer",
         [r"C:\Program Files (x86)\Kingsoft\WPS Office\wps.exe",
          r"C:\Program Files\Kingsoft\WPS Office\wps.exe"]),
        ("OpenOffice Writer",
         [r"C:\Program Files\OpenOffice 4\program\swriter.exe",
          r"C:\Program Files (x86)\OpenOffice 4\program\swriter.exe"]),
        ("WordPad",
         [r"C:\Program Files\Windows NT\Accessories\wordpad.exe",
          r"C:\Windows\System32\write.exe"]),
    ]
    for name, paths in candidates:
        for path in paths:
            if os.path.exists(path):
                editors.append((name, path))
                break
    return editors


def open_with_editor(doc_path, editor_path):
    try:
        import subprocess
        subprocess.Popen([editor_path, doc_path])
    except Exception:
        try:
            os.startfile(doc_path)
        except Exception:
            pass


# ── Main Application ──────────────────────────────────────────────────────────

class PageCaptureApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"PageCapture {APP_VERSION}")
        self.root.geometry("480x400")
        self.root.minsize(440, 360)
        self.root.configure(bg=COLORS['bg'])
        self.root.resizable(True, True)

        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  // 2) - 240
        y = (self.root.winfo_screenheight() // 2) - 200
        self.root.geometry(f"+{x}+{y}")

        self.last_data = None

        # Start Native Messaging listener (replaces old HTTP server)
        start_native_listener(self)

        self._show_waiting()
        self.root.mainloop()

    # ── Waiting screen ────────────────────────────────────────────────────────

    def _show_waiting(self):
        self._clear()
        self.root.geometry("480x400")

        self._header(
            "PageCapture is ready",
            "Click the PageCapture button in Firefox to capture a page")

        body = tk.Frame(self.root, bg=COLORS['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=16)

        steps = [
            ("1", "Open the page you want to capture in Firefox"),
            ("2", "Click the PageCapture button in your Firefox toolbar"),
            ("3", "The page content will appear here automatically"),
            ("4", "Choose your document editor and save"),
        ]

        for num, desc in steps:
            row = tk.Frame(body, bg=COLORS['bg'])
            row.pack(fill='x', pady=6)

            tk.Label(row,
                     text=num,
                     bg=COLORS['purple'],
                     fg=COLORS['bg'],
                     font=('Arial', 11, 'bold'),
                     width=3, height=1).pack(side='left', padx=(0,12))

            tk.Label(row,
                     text=desc,
                     bg=COLORS['bg'],
                     fg=COLORS['text'],
                     font=('Arial', 11),
                     anchor='w',
                     wraplength=360,
                     justify='left').pack(side='left', fill='x', expand=True)

        self.status_var = tk.StringVar(value="Waiting for Firefox extension...")
        tk.Label(body,
                 textvariable=self.status_var,
                 bg=COLORS['bg'],
                 fg=COLORS['blue'],
                 font=('Arial', 10, 'italic')).pack(pady=16)

        tk.Button(body,
                  text="⚙  Change save folder",
                  bg=COLORS['panel'],
                  fg=COLORS['muted'],
                  font=('Arial', 10),
                  relief='flat',
                  cursor='hand2',
                  padx=10, pady=4,
                  command=self._settings).pack()

        self._footer()

    # ── Content received from extension ──────────────────────────────────────

    def on_content_received(self, data):
        self.last_data = data
        title  = data.get('title', 'Captured Page')
        blocks = data.get('blocks', [])

        if not blocks:
            self.status_var.set("No content received. Try again.")
            return

        self._show_save_screen(title, data.get('url', ''), blocks)

    # ── Save screen ───────────────────────────────────────────────────────────

    def _show_save_screen(self, title, url, blocks):
        self._clear()
        self.root.geometry("520x460")

        self._header(
            "Page captured!",
            "Choose how you want to open it")

        body = tk.Frame(self.root, bg=COLORS['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=12)

        tk.Label(body,
                 text="Captured page:",
                 bg=COLORS['bg'],
                 fg=COLORS['muted'],
                 font=('Arial', 10)).pack(anchor='w')

        tk.Label(body,
                 text=title[:70] + ('...' if len(title) > 70 else ''),
                 bg=COLORS['panel'],
                 fg=COLORS['purple'],
                 font=('Arial', 11, 'bold'),
                 padx=12, pady=8,
                 wraplength=440,
                 justify='left').pack(fill='x', pady=(2,12))

        text_count = sum(1 for b in blocks if b.get('type') == 'text')
        img_count  = sum(1 for b in blocks if b.get('type') == 'image')
        link_count = sum(1 for b in blocks if b.get('type') == 'link')

        tk.Label(body,
                 text=f"{len(blocks)} items captured  —  "
                      f"{text_count} text blocks  •  "
                      f"{img_count} images  •  "
                      f"{link_count} links",
                 bg=COLORS['bg'],
                 fg=COLORS['muted'],
                 font=('Arial', 10)).pack(anchor='w', pady=(0,12))

        editors = find_doc_editors()

        if editors:
            tk.Label(body,
                     text="Open with:",
                     bg=COLORS['bg'],
                     fg=COLORS['text'],
                     font=('Arial', 11, 'bold')).pack(anchor='w', pady=(0,6))

            for name, path in editors:
                tk.Button(body,
                          text=f"▶  Open in {name}",
                          bg=COLORS['panel'],
                          fg=COLORS['text'],
                          font=('Arial', 11),
                          relief='flat',
                          cursor='hand2',
                          padx=12, pady=7,
                          anchor='w',
                          command=lambda p=path: self._save_and_open(
                              title, blocks, p)
                          ).pack(fill='x', pady=2)

        tk.Button(body,
                  text="💾  Save to folder only (no editor)",
                  bg=COLORS['panel'],
                  fg=COLORS['muted'],
                  font=('Arial', 10),
                  relief='flat',
                  cursor='hand2',
                  padx=12, pady=6,
                  command=lambda: self._save_and_open(title, blocks, None)
                  ).pack(fill='x', pady=(8,2))

        self.save_status = tk.StringVar(value='')
        tk.Label(body,
                 textvariable=self.save_status,
                 bg=COLORS['bg'],
                 fg=COLORS['green'],
                 font=('Arial', 10)).pack(pady=4)

        self._footer()

    # ── Save and open ─────────────────────────────────────────────────────────

    def _save_and_open(self, title, blocks, editor_path):
        cfg    = load_config()
        folder = cfg.get('save_folder', SAVE_FOLDER)
        os.makedirs(folder, exist_ok=True)

        safe = re.sub(r'[\\/:*?"<>|]', '_', title or 'capture')[:60]
        path = os.path.join(folder, safe + '.docx')

        counter = 1
        while os.path.exists(path):
            path = os.path.join(folder, f"{safe}_{counter}.docx")
            counter += 1

        self.save_status.set("Converting to DOCX...")
        self.root.update()

        def do_save():
            try:
                blocks_to_docx(title, blocks, path)
                self.root.after(0, lambda: self._on_saved(path, editor_path))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda msg=err: self.save_status.set(
                    f"Error: {msg}"))

        threading.Thread(target=do_save, daemon=True).start()

    def _on_saved(self, path, editor_path):
        self.save_status.set(f"Saved: {os.path.basename(path)}")
        if editor_path:
            open_with_editor(path, editor_path)
        else:
            try:
                os.startfile(os.path.dirname(path))
            except Exception:
                pass

        self.root.after(4000, self._show_waiting)

    # ── Settings ──────────────────────────────────────────────────────────────

    def _settings(self):
        cfg = load_config()
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("460x160")
        win.configure(bg=COLORS['bg'])
        win.resizable(False, False)

        tk.Label(win,
                 text="Save folder:",
                 bg=COLORS['bg'],
                 fg=COLORS['text'],
                 font=('Arial', 11)).pack(anchor='w', padx=16, pady=(16,4))

        row = tk.Frame(win, bg=COLORS['bg'])
        row.pack(fill='x', padx=16)

        folder_var = tk.StringVar(value=cfg.get('save_folder', SAVE_FOLDER))
        tk.Entry(row,
                 textvariable=folder_var,
                 bg=COLORS['panel'],
                 fg=COLORS['text'],
                 insertbackground='white',
                 relief='flat',
                 font=('Arial', 10)).pack(
                     side='left', fill='x', expand=True, ipady=5)

        def browse():
            d = filedialog.askdirectory(initialdir=folder_var.get())
            if d:
                folder_var.set(d)

        tk.Button(row,
                  text='Browse',
                  bg=COLORS['border'],
                  fg=COLORS['text'],
                  relief='flat',
                  command=browse,
                  padx=8).pack(side='left', padx=(6,0))

        def do_save():
            cfg['save_folder'] = folder_var.get()
            save_config(cfg)
            win.destroy()

        tk.Button(win,
                  text='Save',
                  bg=COLORS['purple'],
                  fg=COLORS['bg'],
                  font=('Arial', 11, 'bold'),
                  relief='flat',
                  command=do_save,
                  padx=12, pady=6).pack(pady=12)

    # ── UI helpers ────────────────────────────────────────────────────────────

    def _clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    def _header(self, title, subtitle=''):
        hdr = tk.Frame(self.root, bg=COLORS['panel'], pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr,
                 text=title,
                 bg=COLORS['panel'],
                 fg=COLORS['purple'],
                 font=('Arial', 14, 'bold'),
                 padx=16).pack(anchor='w')
        if subtitle:
            tk.Label(hdr,
                     text=subtitle,
                     bg=COLORS['panel'],
                     fg=COLORS['muted'],
                     font=('Arial', 10),
                     padx=16).pack(anchor='w')

    def _footer(self):
        footer = tk.Frame(self.root, bg=COLORS['bg'], pady=8)
        footer.pack(fill='x', side='bottom', padx=16)
        tk.Button(footer,
                  text="✕  Exit PageCapture",
                  bg=COLORS['red'],
                  fg=COLORS['bg'],
                  font=('Arial', 10, 'bold'),
                  relief='flat',
                  cursor='hand2',
                  padx=10, pady=4,
                  command=self.root.quit).pack(side='right')


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    PageCaptureApp()
