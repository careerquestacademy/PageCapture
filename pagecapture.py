"""
PageCapture — Windows desktop app
Captures full content from any browser or chat window,
lets the user select from point A to point B, then saves or copies.
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ctypes
import ctypes.wintypes
import subprocess
import json
import re

# ── Graceful imports ──────────────────────────────────────────────────────────

try:
    import pystray
    from pystray import MenuItem, Menu
    from PIL import Image, ImageDraw, ImageFont
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

try:
    import keyboard
    HAS_HOTKEY = True
except ImportError:
    HAS_HOTKEY = False

try:
    import comtypes.client
    import comtypes.gen
    HAS_UIA = True
except ImportError:
    HAS_UIA = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import pyperclip
    HAS_CLIP = True
except ImportError:
    HAS_CLIP = False

# ── Config ────────────────────────────────────────────────────────────────────

APP_NAME    = "PageCapture"
HOTKEY      = "win+shift+c"
SAVE_FOLDER = os.path.expanduser("~/Documents/PageCapture")
os.makedirs(SAVE_FOLDER, exist_ok=True)

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".pagecapture_config.json")

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

# ── Windows Accessibility capture ────────────────────────────────────────────

def get_foreground_hwnd():
    user32 = ctypes.windll.user32
    return user32.GetForegroundWindow()

def get_window_title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    ctypes.windll.user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value

def capture_via_uia(hwnd):
    """
    Use Windows UI Automation to extract full text + links + image alts
    from any accessible window (browsers, Electron apps, etc.)
    Returns list of content blocks.
    """
    blocks = []
    try:
        import comtypes.client
        uia = comtypes.client.CreateObject(
            "{ff48dba4-60ef-4201-aa87-54103eef594e}",
            interface=comtypes.gen.UIAutomationClient.IUIAutomation
        )
        element = uia.ElementFromHandle(hwnd)
        if not element:
            return blocks

        # Walk all text ranges
        text_pattern_id = 10014  # UIA_TextPatternId
        try:
            tp = element.GetCurrentPattern(text_pattern_id)
            if tp:
                from comtypes.gen.UIAutomationClient import IUIAutomationTextPattern
                itp = tp.QueryInterface(IUIAutomationTextPattern)
                doc_range = itp.DocumentRange
                text = doc_range.GetText(-1)
                if text and text.strip():
                    # Split into paragraphs
                    for line in text.split('\n'):
                        line = line.strip()
                        if line:
                            blocks.append({'type': 'text', 'text': line})
                    return blocks
        except Exception:
            pass

        # Fallback: walk element tree
        walker = uia.CreateTreeWalker(uia.ControlViewCondition)
        _walk_element(uia, walker, element, blocks, depth=0)

    except Exception as e:
        blocks.append({'type': 'text', 'text': f'[Capture error: {str(e)}]'})

    return blocks


def _walk_element(uia, walker, element, blocks, depth):
    if depth > 30:
        return
    try:
        ctrl_type = element.CurrentControlType
        name      = (element.CurrentName or '').strip()
        value     = ''

        # Try value pattern
        try:
            vp = element.GetCurrentPattern(10002)  # UIA_ValuePatternId
            if vp:
                from comtypes.gen.UIAutomationClient import IUIAutomationValuePattern
                ivp = vp.QueryInterface(IUIAutomationValuePattern)
                value = (ivp.CurrentValue or '').strip()
        except Exception:
            pass

        # Control type constants
        # 50020 = Hyperlink, 50031 = Image, 50004 = Button
        # 50025 = Text, 50033 = Edit, 50032 = Document
        # 50021 = List, 50023 = ListItem, 50026 = Header

        if ctrl_type == 50020 and name:  # Hyperlink
            # Try to get URL
            try:
                legacy = element.GetCurrentPattern(10018)
                from comtypes.gen.UIAutomationClient import IUIAutomationLegacyIAccessiblePattern
                il = legacy.QueryInterface(IUIAutomationLegacyIAccessiblePattern)
                href = il.CurrentValue or ''
            except Exception:
                href = ''
            blocks.append({'type': 'link', 'text': name, 'href': href})
            return

        if ctrl_type == 50031:  # Image
            alt = name or value or '[image]'
            blocks.append({'type': 'image', 'alt': alt})
            return

        if ctrl_type in (50025, 50033) and (name or value):  # Text, Edit
            content = name or value
            if content and len(content) > 1:
                blocks.append({'type': 'text', 'text': content})

        if ctrl_type == 50023 and name:  # ListItem
            blocks.append({'type': 'listitem', 'text': name})

        # Headings — check name patterns
        if ctrl_type == 50025 and name:
            for lvl in range(1, 7):
                role_name = f'heading level {lvl}'
                try:
                    lp = element.GetCurrentPattern(10018)
                    from comtypes.gen.UIAutomationClient import IUIAutomationLegacyIAccessiblePattern
                    il = lp.QueryInterface(IUIAutomationLegacyIAccessiblePattern)
                    role = il.CurrentRole
                    if role == 10:  # ROLE_SYSTEM_GROUPING / heading
                        blocks.append({'type': 'heading', 'level': 2, 'text': name})
                        return
                except Exception:
                    pass

    except Exception:
        pass

    # Recurse children
    try:
        child = walker.GetFirstChildElement(element)
        while child:
            _walk_element(uia, walker, child, blocks, depth + 1)
            try:
                child = walker.GetNextSiblingElement(child)
            except Exception:
                break
    except Exception:
        pass


def capture_via_clipboard(hwnd):
    """
    Fallback: send Ctrl+A, Ctrl+C to the window,
    read clipboard text, then restore clipboard.
    """
    blocks = []
    try:
        import win32gui
        import win32con
        import win32api
        import time

        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.15)

        # Save clipboard
        if HAS_CLIP:
            try:
                saved = pyperclip.paste()
            except Exception:
                saved = ''

        # Select all + copy
        win32api.keybd_event(0x11, 0, 0, 0)           # Ctrl down
        win32api.keybd_event(0x41, 0, 0, 0)           # A
        win32api.keybd_event(0x41, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)
        win32api.keybd_event(0x11, 0, 0, 0)           # Ctrl down
        win32api.keybd_event(0x43, 0, 0, 0)           # C
        win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.2)

        if HAS_CLIP:
            text = pyperclip.paste()
            # Restore clipboard
            try:
                pyperclip.copy(saved)
            except Exception:
                pass
        else:
            text = ''

        for line in text.split('\n'):
            line = line.strip()
            if line:
                blocks.append({'type': 'text', 'text': line})

    except Exception as e:
        blocks.append({'type': 'text', 'text': f'[Fallback capture error: {str(e)}]'})

    return blocks


def capture_window():
    """
    Main capture entry point.
    Tries UIA first, falls back to clipboard method.
    """
    hwnd  = get_foreground_hwnd()
    title = get_window_title(hwnd)

    blocks = []
    if HAS_UIA:
        blocks = capture_via_uia(hwnd)

    if not blocks:
        blocks = capture_via_clipboard(hwnd)

    if not blocks:
        blocks = [{'type': 'text', 'text': '[Nothing could be captured from this window.]'}]

    return title, blocks


# ── Reader Window ─────────────────────────────────────────────────────────────

class ReaderWindow:
    def __init__(self, title, blocks):
        self.title  = title
        self.blocks = blocks
        self.root   = tk.Tk()
        self.root.title(f"PageCapture — {title}")
        self.root.geometry("820x640")
        self.root.minsize(600, 400)
        self.root.configure(bg='#1e1e2e')

        # Track selection
        self.sel_start = None
        self.sel_end   = None
        self.click_count = 0

        self._build_ui()
        self._populate()
        self.root.mainloop()

    def _build_ui(self):
        # ── Toolbar ──
        toolbar = tk.Frame(self.root, bg='#313244', pady=6)
        toolbar.pack(fill='x', side='top')

        tk.Label(toolbar, text="PageCapture",
                 bg='#313244', fg='#cba6f7',
                 font=('Arial', 13, 'bold')).pack(side='left', padx=12)

        tk.Label(toolbar, text="Click once to set START  •  Click again to set END",
                 bg='#313244', fg='#a6adc8',
                 font=('Arial', 10)).pack(side='left', padx=8)

        # Right-side buttons
        btn_cfg = dict(bg='#45475a', fg='#cdd6f4', relief='flat',
                       font=('Arial', 10, 'bold'), padx=10, pady=4,
                       cursor='hand2', activebackground='#585b70',
                       activeforeground='white')

        tk.Button(toolbar, text='✕ Clear',
                  command=self._clear_selection, **btn_cfg).pack(side='right', padx=4)
        tk.Button(toolbar, text='📋 Copy',
                  command=self._copy_selection, **btn_cfg).pack(side='right', padx=4)
        tk.Button(toolbar, text='💾 TXT',
                  command=lambda: self._save('txt'), **btn_cfg).pack(side='right', padx=4)
        tk.Button(toolbar, text='💾 HTML',
                  command=lambda: self._save('html'), **btn_cfg).pack(side='right', padx=4)
        tk.Button(toolbar, text='💾 DOCX',
                  command=lambda: self._save('docx'), **btn_cfg).pack(side='right', padx=4)

        # ── Status bar ──
        self.status_var = tk.StringVar(value="Click anywhere in the text to set your START point.")
        status = tk.Label(self.root, textvariable=self.status_var,
                          bg='#181825', fg='#89b4fa',
                          font=('Arial', 10), anchor='w', padx=10, pady=4)
        status.pack(fill='x', side='bottom')

        # ── Text area ──
        frame = tk.Frame(self.root, bg='#1e1e2e')
        frame.pack(fill='both', expand=True, padx=0, pady=0)

        scrollbar = tk.Scrollbar(frame, bg='#313244', troughcolor='#1e1e2e',
                                 activebackground='#cba6f7')
        scrollbar.pack(side='right', fill='y')

        self.text = tk.Text(
            frame,
            wrap='word',
            bg='#1e1e2e',
            fg='#cdd6f4',
            font=('Georgia', 12),
            padx=20, pady=16,
            spacing1=4, spacing3=4,
            cursor='arrow',
            yscrollcommand=scrollbar.set,
            relief='flat',
            borderwidth=0,
            selectbackground='#cba6f7',
            selectforeground='#1e1e2e',
            insertwidth=0
        )
        self.text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.text.yview)

        # Tags
        self.text.tag_config('heading1', font=('Arial', 18, 'bold'), foreground='#cba6f7', spacing1=12, spacing3=6)
        self.text.tag_config('heading2', font=('Arial', 15, 'bold'), foreground='#cba6f7', spacing1=10, spacing3=4)
        self.text.tag_config('heading3', font=('Arial', 13, 'bold'), foreground='#b4befe', spacing1=8,  spacing3=4)
        self.text.tag_config('link',     font=('Georgia', 12, 'underline'), foreground='#89b4fa')
        self.text.tag_config('listitem', foreground='#a6e3a1', lmargin1=24, lmargin2=24)
        self.text.tag_config('image',    foreground='#f9e2af', font=('Arial', 10, 'italic'))
        self.text.tag_config('selected', background='#45475a', foreground='#ffffff')
        self.text.tag_config('marker_start', background='#a6e3a1', foreground='#1e1e2e')
        self.text.tag_config('marker_end',   background='#f38ba8', foreground='#1e1e2e')

        # Click to select
        self.text.bind('<Button-1>', self._on_click)
        self.text.config(state='disabled')

    def _populate(self):
        self.text.config(state='normal')
        self.text.delete('1.0', 'end')
        self.line_map = []  # maps line index to block index

        for i, block in enumerate(self.blocks):
            btype = block.get('type', 'text')
            text  = block.get('text', '')
            level = block.get('level', 1)
            href  = block.get('href', '')
            alt   = block.get('alt', '')

            start = self.text.index('end')

            if btype == 'heading':
                tag = f'heading{min(level, 3)}'
                self.text.insert('end', text + '\n', tag)
            elif btype == 'link':
                display = text + (f'  [{href}]' if href else '')
                self.text.insert('end', display + '\n', 'link')
            elif btype == 'listitem':
                self.text.insert('end', '  • ' + text + '\n', 'listitem')
            elif btype == 'image':
                self.text.insert('end', f'[Image: {alt}]\n', 'image')
            else:
                self.text.insert('end', text + '\n')

            end = self.text.index('end')
            self.line_map.append((start, end, i))

        self.text.config(state='disabled')

    def _on_click(self, event):
        idx = self.text.index(f'@{event.x},{event.y}')

        if self.click_count == 0:
            self.sel_start = idx
            self.click_count = 1
            self._highlight_selection()
            self.status_var.set(f"START set at line {idx.split('.')[0]}  —  Now click to set your END point.")
        else:
            self.sel_end = idx
            self.click_count = 0
            self._highlight_selection()
            # Count approximate lines selected
            l1 = int(self.sel_start.split('.')[0])
            l2 = int(self.sel_end.split('.')[0])
            if l2 < l1:
                self.sel_start, self.sel_end = self.sel_end, self.sel_start
                l1, l2 = l2, l1
            count = l2 - l1
            self.status_var.set(f"Selected {count} lines  —  Use the buttons above to Copy or Save.")

    def _highlight_selection(self):
        self.text.config(state='normal')
        self.text.tag_remove('selected', '1.0', 'end')
        if self.sel_start and self.sel_end:
            s = self.sel_start
            e = self.sel_end
            if self.text.compare(s, '>', e):
                s, e = e, s
            self.text.tag_add('selected', s, e)
        self.text.config(state='disabled')

    def _get_selected_blocks(self):
        if not self.sel_start or not self.sel_end:
            return self.blocks  # no selection = use everything

        s = self.sel_start
        e = self.sel_end
        if self.text.compare(s, '>', e):
            s, e = e, s

        selected = []
        for (bstart, bend, idx) in self.line_map:
            # Include block if it overlaps the selection range
            if (self.text.compare(bstart, '<=', e) and
                self.text.compare(bend,   '>=', s)):
                selected.append(self.blocks[idx])

        return selected if selected else self.blocks

    def _get_selected_text(self):
        blocks = self._get_selected_blocks()
        lines = []
        for b in blocks:
            btype = b.get('type', 'text')
            if btype == 'heading':
                lines.append('\n' + b.get('text','').upper() + '\n')
            elif btype == 'listitem':
                lines.append('  • ' + b.get('text',''))
            elif btype == 'link':
                href = b.get('href','')
                text = b.get('text','')
                lines.append(text + (f'  [{href}]' if href else ''))
            elif btype == 'image':
                lines.append(f"[Image: {b.get('alt','')}]")
            else:
                lines.append(b.get('text',''))
        return '\n'.join(lines)

    def _copy_selection(self):
        text = self._get_selected_text()
        if not text.strip():
            self.status_var.set("Nothing to copy.")
            return
        if HAS_CLIP:
            pyperclip.copy(text)
        else:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        self.status_var.set("Copied to clipboard!")

    def _clear_selection(self):
        self.sel_start  = None
        self.sel_end    = None
        self.click_count = 0
        self.text.config(state='normal')
        self.text.tag_remove('selected', '1.0', 'end')
        self.text.config(state='disabled')
        self.status_var.set("Selection cleared. Click to set a new START point.")

    def _save(self, fmt):
        cfg     = load_config()
        folder  = cfg.get('save_folder', SAVE_FOLDER)
        os.makedirs(folder, exist_ok=True)

        safe    = re.sub(r'[\\/:*?"<>|]', '_', self.title or 'capture')[:60]
        blocks  = self._get_selected_blocks()

        if fmt == 'txt':
            path = os.path.join(folder, safe + '.txt')
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self._get_selected_text())
            self.status_var.set(f"Saved: {path}")

        elif fmt == 'html':
            path = os.path.join(folder, safe + '.html')
            _save_html(path, self.title, blocks)
            self.status_var.set(f"Saved: {path}")

        elif fmt == 'docx':
            if not HAS_DOCX:
                self.status_var.set("python-docx not available. Use TXT or HTML.")
                return
            path = os.path.join(folder, safe + '.docx')
            _save_docx(path, self.title, blocks)
            self.status_var.set(f"Saved: {path}")

        # Open the folder so user can find it
        try:
            os.startfile(folder)
        except Exception:
            pass

    def _clear_selection(self):
        self.sel_start   = None
        self.sel_end     = None
        self.click_count = 0
        self.text.config(state='normal')
        self.text.tag_remove('selected', '1.0', 'end')
        self.text.config(state='disabled')
        self.status_var.set("Selection cleared. Click to set a new START point.")


# ── Save helpers ──────────────────────────────────────────────────────────────

def _save_html(path, title, blocks):
    lines = []
    for b in blocks:
        btype = b.get('type','text')
        if btype == 'heading':
            lvl  = b.get('level',2)
            text = _esc(b.get('text',''))
            lines.append(f'<h{lvl}>{text}</h{lvl}>')
        elif btype == 'link':
            text = _esc(b.get('text',''))
            href = _esc(b.get('href',''))
            lines.append(f'<p><a href="{href}">{text}</a></p>')
        elif btype == 'listitem':
            lines.append(f'<li>{_esc(b.get("text",""))}</li>')
        elif btype == 'image':
            lines.append(f'<p><em>[Image: {_esc(b.get("alt",""))}]</em></p>')
        else:
            lines.append(f'<p>{_esc(b.get("text",""))}</p>')

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>{_esc(title)}</title>
  <style>
    body {{ font-family: Georgia, serif; max-width: 820px; margin: 40px auto;
            padding: 0 20px; line-height: 1.75; color: #222; }}
    h1,h2,h3 {{ margin-top: 1.4em; color: #333; }}
    a  {{ color: #0066cc; }}
    li {{ margin: 0.3em 0; }}
    p  {{ margin: 0.5em 0; }}
  </style>
</head>
<body>
<h1>{_esc(title)}</h1>
{''.join(lines)}
</body>
</html>"""

    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)


def _save_docx(path, title, blocks):
    doc = Document()

    # Title
    title_para = doc.add_paragraph()
    run = title_para.add_run(title or 'Capture')
    run.bold = True
    run.font.size = Pt(18)

    for b in blocks:
        btype = b.get('type','text')
        if btype == 'heading':
            lvl  = min(b.get('level',2), 9)
            doc.add_heading(b.get('text',''), level=lvl)
        elif btype == 'link':
            p   = doc.add_paragraph()
            run = p.add_run(b.get('text',''))
            run.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
            run.underline = True
            href = b.get('href','')
            if href:
                p.add_run(f'  [{href}]')
        elif btype == 'listitem':
            doc.add_paragraph(b.get('text',''), style='List Bullet')
        elif btype == 'image':
            p   = doc.add_paragraph()
            run = p.add_run(f"[Image: {b.get('alt','')}]")
            run.italic = True
        else:
            doc.add_paragraph(b.get('text',''))

    doc.save(path)


def _esc(s):
    return str(s or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')


# ── System tray ───────────────────────────────────────────────────────────────

def make_tray_icon():
    img  = Image.new('RGBA', (64, 64), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([4,4,60,60], radius=12, fill=(203,166,247,255))
    draw.rectangle([18,16,22,48], fill='white')
    draw.rectangle([18,16,46,20], fill='white')
    draw.rectangle([18,32,46,36], fill='white')
    draw.rectangle([42,16,46,36], fill='white')
    return img


def do_capture():
    """Called from hotkey or tray — runs capture in background thread."""
    def _run():
        import time
        time.sleep(0.3)  # let hotkey window lose focus
        title, blocks = capture_window()
        # Open reader on main thread
        ReaderWindow(title, blocks)

    t = threading.Thread(target=_run, daemon=True)
    t.start()


def start_tray():
    if not HAS_TRAY:
        return

    icon_img = make_tray_icon()

    def on_capture(icon, item):
        do_capture()

    def on_settings(icon, item):
        _open_settings()

    def on_quit(icon, item):
        icon.stop()
        os._exit(0)

    menu = Menu(
        MenuItem('Capture Window  (Win+Shift+C)', on_capture, default=True),
        MenuItem('Settings', on_settings),
        Menu.SEPARATOR,
        MenuItem('Quit PageCapture', on_quit)
    )

    icon = pystray.Icon(APP_NAME, icon_img, APP_NAME, menu)

    if HAS_HOTKEY:
        keyboard.add_hotkey(HOTKEY, do_capture, suppress=True)

    icon.run()


def _open_settings():
    cfg  = load_config()
    root = tk.Tk()
    root.title("PageCapture Settings")
    root.geometry("460x160")
    root.configure(bg='#1e1e2e')
    root.resizable(False, False)

    tk.Label(root, text="Save folder:", bg='#1e1e2e', fg='#cdd6f4',
             font=('Arial',11)).pack(anchor='w', padx=16, pady=(16,4))

    row = tk.Frame(root, bg='#1e1e2e')
    row.pack(fill='x', padx=16)

    folder_var = tk.StringVar(value=cfg.get('save_folder', SAVE_FOLDER))
    entry = tk.Entry(row, textvariable=folder_var, bg='#313244', fg='#cdd6f4',
                     insertbackground='white', relief='flat', font=('Arial',10))
    entry.pack(side='left', fill='x', expand=True, ipady=5)

    def browse():
        d = filedialog.askdirectory(initialdir=folder_var.get())
        if d:
            folder_var.set(d)

    tk.Button(row, text='Browse', bg='#45475a', fg='#cdd6f4', relief='flat',
              command=browse, padx=8).pack(side='left', padx=(6,0))

    def save():
        cfg['save_folder'] = folder_var.get()
        save_config(cfg)
        root.destroy()

    tk.Button(root, text='Save Settings', bg='#cba6f7', fg='#1e1e2e',
              font=('Arial',11,'bold'), relief='flat', command=save,
              padx=12, pady=6).pack(pady=16)

    root.mainloop()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # If run directly without tray support, just open capture immediately
    if not HAS_TRAY:
        title, blocks = capture_window()
        ReaderWindow(title, blocks)
    else:
        start_tray()
