"""
PageCapture v2
A friendly, safe tool for capturing text from browsers and documents.
NO keyboard hooks, NO mouse hooks, NO system-level anything.
Just a plain window that minds its own business.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import re
import json
import time

# ── Graceful imports ──────────────────────────────────────────────────────────

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import pyperclip
    HAS_CLIP = True
except ImportError:
    HAS_CLIP = False

try:
    import win32gui
    import win32con
    import win32api
    import win32process
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import comtypes.client
    HAS_UIA = True
except ImportError:
    HAS_UIA = False

# ── Config ────────────────────────────────────────────────────────────────────

APP_NAME    = "PageCapture"
APP_VERSION = "2.0"
SAVE_FOLDER = os.path.expanduser("~/Documents/PageCapture")
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".pagecapture2_config.json")

os.makedirs(SAVE_FOLDER, exist_ok=True)

COLORS = {
    'bg':       '#1e1e2e',
    'panel':    '#313244',
    'border':   '#45475a',
    'text':     '#cdd6f4',
    'muted':    '#a6adc8',
    'purple':   '#cba6f7',
    'blue':     '#89b4fa',
    'green':    '#a6e3a1',
    'red':      '#f38ba8',
    'yellow':   '#f9e2af',
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

# ── Window enumeration ────────────────────────────────────────────────────────

def get_open_windows():
    """Get list of visible, named windows the user can capture from."""
    windows = []
    if not HAS_WIN32:
        return windows

    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title or len(title) < 2:
            return
        # Skip system windows
        skip = ['Program Manager', 'Windows Input Experience',
                'PageCapture', 'Task Switching']
        if any(s in title for s in skip):
            return
        windows.append((hwnd, title))

    win32gui.EnumWindows(callback, None)
    return windows

def get_browser_tabs():
    """Get all open browser windows."""
    all_windows = get_open_windows()
    browsers = []
    browser_names = ['firefox', 'chrome', 'edge', 'opera', 'brave', 'safari']
    for hwnd, title in all_windows:
        tl = title.lower()
        if any(b in tl for b in browser_names):
            browsers.append((hwnd, title))
    return browsers

def get_document_windows():
    """Get open document/chat windows."""
    all_windows = get_open_windows()
    docs = []
    browser_names = ['firefox', 'chrome', 'edge', 'opera', 'brave']
    for hwnd, title in all_windows:
        tl = title.lower()
        if not any(b in tl for b in browser_names):
            docs.append((hwnd, title))
    return docs

# ── Content capture ───────────────────────────────────────────────────────────

def capture_window_content(hwnd):
    """
    Safely capture text content from a window.
    Uses UIA accessibility API — no keyboard hooks, no mouse hooks.
    Falls back to a safe clipboard method if UIA fails.
    """
    blocks = []

    # Try UIA first
    if HAS_UIA:
        try:
            blocks = _capture_uia(hwnd)
        except Exception:
            blocks = []

    # Safe clipboard fallback
    if not blocks and HAS_WIN32:
        try:
            blocks = _capture_clipboard_safe(hwnd)
        except Exception:
            blocks = []

    if not blocks:
        blocks = [{'type': 'text',
                   'text': 'Could not capture content from this window. '
                           'Try a different window or source type.'}]
    return blocks


def _capture_uia(hwnd):
    """Capture via Windows UI Automation — read only, no input simulation."""
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

        try:
            tp = element.GetCurrentPattern(10014)
            if tp:
                from comtypes.gen.UIAutomationClient import IUIAutomationTextPattern
                itp = tp.QueryInterface(IUIAutomationTextPattern)
                text = itp.DocumentRange.GetText(-1)
                if text and text.strip():
                    for line in text.split('\n'):
                        line = line.strip()
                        if line:
                            blocks.append({'type': 'text', 'text': line})
                    return blocks
        except Exception:
            pass

    except Exception:
        pass
    return blocks


def _capture_clipboard_safe(hwnd):
    """
    Safe clipboard capture — brings window to front, uses Ctrl+A Ctrl+C,
    then immediately restores focus. No suppression, no hooks.
    """
    blocks = []
    try:
        # Save current clipboard content
        saved_clip = ''
        if HAS_CLIP:
            try:
                saved_clip = pyperclip.paste()
            except Exception:
                pass

        # Bring target window to front
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)

        # Select all and copy using PostMessage — safer than keybd_event
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, 0x41, 0)  # A key
        time.sleep(0.05)

        # Use SendKeys via shell for Ctrl+A, Ctrl+C
        import subprocess
        script = f"""
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait('^a')
Start-Sleep -Milliseconds 100
[System.Windows.Forms.SendKeys]::SendWait('^c')
Start-Sleep -Milliseconds 200
"""
        subprocess.run(['powershell', '-Command', script],
                      capture_output=True, timeout=5)

        time.sleep(0.3)

        if HAS_CLIP:
            text = pyperclip.paste()
            # Restore clipboard
            try:
                pyperclip.copy(saved_clip)
            except Exception:
                pass

            if text and text.strip() and text != saved_clip:
                for line in text.split('\n'):
                    line = line.strip()
                    if line:
                        blocks.append({'type': 'text', 'text': line})

    except Exception as e:
        blocks.append({'type': 'text',
                       'text': f'Capture note: {str(e)}'})
    return blocks

# ── Main Application Window ───────────────────────────────────────────────────

class PageCaptureApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"PageCapture {APP_VERSION}")
        self.root.geometry("500x420")
        self.root.minsize(480, 380)
        self.root.configure(bg=COLORS['bg'])
        self.root.resizable(True, True)

        # Center on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  // 2) - 250
        y = (self.root.winfo_screenheight() // 2) - 210
        self.root.geometry(f"+{x}+{y}")

        self.selected_hwnd  = None
        self.selected_title = ''
        self.captured_blocks = []

        self._show_welcome()
        self.root.mainloop()

    # ── Step 1: Welcome ───────────────────────────────────────────────────────

    def _show_welcome(self):
        self._clear()

        self._header("Welcome to PageCapture",
                     "Copy text from any browser, chat, or document")

        body = tk.Frame(self.root, bg=COLORS['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=8)

        tk.Label(body,
                 text="What would you like to copy from?",
                 bg=COLORS['bg'], fg=COLORS['text'],
                 font=('Arial', 12)).pack(anchor='w', pady=(8,12))

        options = [
            ("🌐  Browser", "browser",
             "Firefox, Chrome, Edge, or any web browser"),
            ("💬  Chat or App", "chat",
             "Discord, Slack, Teams, or any other app"),
            ("📄  Document", "document",
             "Word, Notepad, PDF viewer, or similar"),
        ]

        for label, key, desc in options:
            btn_frame = tk.Frame(body, bg=COLORS['panel'],
                                 cursor='hand2')
            btn_frame.pack(fill='x', pady=4)
            btn_frame.bind('<Enter>',
                lambda e, f=btn_frame: f.configure(bg=COLORS['border']))
            btn_frame.bind('<Leave>',
                lambda e, f=btn_frame: f.configure(bg=COLORS['panel']))

            tk.Label(btn_frame, text=label,
                     bg=COLORS['panel'], fg=COLORS['purple'],
                     font=('Arial', 12, 'bold'),
                     padx=16, pady=10).pack(side='left')
            tk.Label(btn_frame, text=desc,
                     bg=COLORS['panel'], fg=COLORS['muted'],
                     font=('Arial', 10)).pack(side='left', padx=8)

            btn_frame.bind('<Button-1>',
                lambda e, k=key: self._show_pick_window(k))
            for child in btn_frame.winfo_children():
                child.bind('<Button-1>',
                    lambda e, k=key: self._show_pick_window(k))
                child.bind('<Enter>',
                    lambda e, f=btn_frame: f.configure(bg=COLORS['border']))
                child.bind('<Leave>',
                    lambda e, f=btn_frame: f.configure(bg=COLORS['panel']))

        self._footer(None, None)

    # ── Step 2: Pick a window ─────────────────────────────────────────────────

    def _show_pick_window(self, source_type):
        self._clear()

        titles = {
            'browser':  'Choose your browser tab',
            'chat':     'Choose your chat or app window',
            'document': 'Choose your document window',
        }
        self._header(titles.get(source_type, 'Choose a window'),
                     "Select the window you want to copy from")

        body = tk.Frame(self.root, bg=COLORS['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=8)

        # Get windows
        if source_type == 'browser':
            windows = get_browser_tabs()
            if not windows:
                windows = get_open_windows()
        elif source_type == 'document':
            windows = get_document_windows()
        else:
            windows = get_open_windows()

        if not windows:
            tk.Label(body,
                     text="No open windows found.\nPlease open your browser or app first, then try again.",
                     bg=COLORS['bg'], fg=COLORS['yellow'],
                     font=('Arial', 11), justify='center').pack(pady=20)
            self._footer(self._show_welcome, None)
            return

        tk.Label(body, text="Click the window you want:",
                 bg=COLORS['bg'], fg=COLORS['text'],
                 font=('Arial', 11)).pack(anchor='w', pady=(4,8))

        # Scrollable list
        list_frame = tk.Frame(body, bg=COLORS['bg'])
        list_frame.pack(fill='both', expand=True)

        scrollbar = tk.Scrollbar(list_frame, bg=COLORS['panel'])
        scrollbar.pack(side='right', fill='y')

        listbox = tk.Listbox(list_frame,
                             bg=COLORS['panel'],
                             fg=COLORS['text'],
                             selectbackground=COLORS['purple'],
                             selectforeground=COLORS['bg'],
                             font=('Arial', 11),
                             relief='flat',
                             borderwidth=0,
                             activestyle='none',
                             yscrollcommand=scrollbar.set)
        listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=listbox.yview)

        for hwnd, title in windows:
            listbox.insert('end', f"  {title}")

        def on_select(e=None):
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            self.selected_hwnd  = windows[idx][0]
            self.selected_title = windows[idx][1]
            self._show_confirm_window()

        listbox.bind('<Double-Button-1>', on_select)
        listbox.bind('<Return>', on_select)

        tk.Button(body,
                  text="Select this window →",
                  bg=COLORS['purple'], fg=COLORS['bg'],
                  font=('Arial', 11, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=12, pady=6,
                  command=on_select).pack(pady=8)

        self._footer(self._show_welcome, None)

    # ── Step 3: Confirm ───────────────────────────────────────────────────────

    def _show_confirm_window(self):
        self._clear()

        self._header("Is this the right window?",
                     "Make sure your browser or app is on the correct page")

        body = tk.Frame(self.root, bg=COLORS['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=8)

        tk.Label(body, text="Selected window:",
                 bg=COLORS['bg'], fg=COLORS['muted'],
                 font=('Arial', 10)).pack(anchor='w', pady=(8,2))

        tk.Label(body, text=self.selected_title,
                 bg=COLORS['panel'], fg=COLORS['purple'],
                 font=('Arial', 12, 'bold'),
                 padx=12, pady=10,
                 wraplength=420,
                 justify='left').pack(fill='x', pady=(0,12))

        tk.Label(body,
                 text="Before continuing:\n\n"
                      "1. Switch to that window now\n"
                      "2. Navigate to the page or section you want\n"
                      "3. Come back here and click Continue",
                 bg=COLORS['bg'], fg=COLORS['text'],
                 font=('Arial', 11),
                 justify='left').pack(anchor='w', pady=8)

        tk.Button(body,
                  text="✓  Yes, continue to capture",
                  bg=COLORS['green'], fg=COLORS['bg'],
                  font=('Arial', 12, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=12, pady=8,
                  command=self._show_capture_instructions).pack(fill='x', pady=8)

        tk.Button(body,
                  text="← Choose a different window",
                  bg=COLORS['panel'], fg=COLORS['text'],
                  font=('Arial', 10),
                  relief='flat', cursor='hand2',
                  padx=12, pady=6,
                  command=self._show_welcome).pack(fill='x')

        self._footer(self._show_welcome, None)

    # ── Step 4: Capture instructions ─────────────────────────────────────────

    def _show_capture_instructions(self):
        self._clear()

        self._header("Ready to capture",
                     "Follow these steps to select what you want")

        body = tk.Frame(self.root, bg=COLORS['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=8)

        tk.Label(body,
                 text="How to select your content:",
                 bg=COLORS['bg'], fg=COLORS['text'],
                 font=('Arial', 11, 'bold')).pack(anchor='w', pady=(8,6))

        steps = [
            ("Step 1", "Click Capture below — PageCapture will read the full content of your selected window"),
            ("Step 2", "The content opens in a reader window"),
            ("Step 3", "Scroll through it and click once to set your START point"),
            ("Step 4", "Scroll to where you want to stop and click again to set your END point"),
            ("Step 5", "Click Copy or Save — done!"),
        ]

        for title, desc in steps:
            row = tk.Frame(body, bg=COLORS['bg'])
            row.pack(fill='x', pady=3)
            tk.Label(row, text=title,
                     bg=COLORS['bg'], fg=COLORS['purple'],
                     font=('Arial', 10, 'bold'),
                     width=8, anchor='w').pack(side='left')
            tk.Label(row, text=desc,
                     bg=COLORS['bg'], fg=COLORS['text'],
                     font=('Arial', 10),
                     wraplength=340, justify='left').pack(side='left')

        self.status_var = tk.StringVar(value='')
        tk.Label(body, textvariable=self.status_var,
                 bg=COLORS['bg'], fg=COLORS['blue'],
                 font=('Arial', 10)).pack(pady=4)

        self.capture_btn = tk.Button(body,
                  text="▶  Capture Now",
                  bg=COLORS['purple'], fg=COLORS['bg'],
                  font=('Arial', 13, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=12, pady=10,
                  command=self._do_capture)
        self.capture_btn.pack(fill='x', pady=8)

        self._footer(self._show_confirm_window, None)

    # ── Do the actual capture ─────────────────────────────────────────────────

    def _do_capture(self):
        self.capture_btn.config(state='disabled', text='Capturing...')
        self.status_var.set('Reading content from window...')

        def run():
            blocks = capture_window_content(self.selected_hwnd)
            self.captured_blocks = blocks
            self.root.after(0, self._show_reader)

        threading.Thread(target=run, daemon=True).start()

    # ── Step 5: Reader window ─────────────────────────────────────────────────

    def _show_reader(self):
        self._clear()
        self.root.geometry("700x580")

        self._header(f"Captured: {self.selected_title[:50]}",
                     "Click once for START • Scroll • Click again for END")

        # Toolbar
        toolbar = tk.Frame(self.root, bg=COLORS['panel'], pady=6)
        toolbar.pack(fill='x')

        btn_cfg = dict(bg=COLORS['border'], fg=COLORS['text'],
                       relief='flat', font=('Arial', 10, 'bold'),
                       padx=10, pady=4, cursor='hand2',
                       activebackground=COLORS['purple'],
                       activeforeground=COLORS['bg'])

        tk.Button(toolbar, text='📋 Copy',
                  command=self._copy_selection, **btn_cfg).pack(side='left', padx=4)
        tk.Button(toolbar, text='💾 DOCX',
                  command=lambda: self._save('docx'), **btn_cfg).pack(side='left', padx=2)
        tk.Button(toolbar, text='💾 HTML',
                  command=lambda: self._save('html'), **btn_cfg).pack(side='left', padx=2)
        tk.Button(toolbar, text='💾 TXT',
                  command=lambda: self._save('txt'), **btn_cfg).pack(side='left', padx=2)
        tk.Button(toolbar, text='✕ Clear Selection',
                  command=self._clear_selection, **btn_cfg).pack(side='left', padx=2)
        tk.Button(toolbar, text='← Start Over',
                  command=self._start_over,
                  bg=COLORS['panel'], fg=COLORS['muted'],
                  relief='flat', font=('Arial', 10),
                  padx=10, pady=4,
                  cursor='hand2').pack(side='right', padx=4)

        # Status
        self.status_var = tk.StringVar(
            value="Click anywhere in the text to set your START point.")
        tk.Label(self.root, textvariable=self.status_var,
                 bg=COLORS['bg'], fg=COLORS['blue'],
                 font=('Arial', 10), anchor='w',
                 padx=12, pady=4).pack(fill='x')

        # Text area
        frame = tk.Frame(self.root, bg=COLORS['bg'])
        frame.pack(fill='both', expand=True)

        scrollbar = tk.Scrollbar(frame, bg=COLORS['panel'],
                                 troughcolor=COLORS['bg'])
        scrollbar.pack(side='right', fill='y')

        self.text = tk.Text(
            frame,
            wrap='word',
            bg=COLORS['bg'],
            fg=COLORS['text'],
            font=('Georgia', 12),
            padx=20, pady=16,
            spacing1=4, spacing3=4,
            cursor='arrow',
            yscrollcommand=scrollbar.set,
            relief='flat',
            borderwidth=0,
            selectbackground=COLORS['purple'],
            selectforeground=COLORS['bg'],
            insertwidth=0
        )
        self.text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.text.yview)

        # Tags
        self.text.tag_config('heading',
            font=('Arial', 14, 'bold'),
            foreground=COLORS['purple'],
            spacing1=10, spacing3=6)
        self.text.tag_config('link',
            font=('Georgia', 12, 'underline'),
            foreground=COLORS['blue'])
        self.text.tag_config('listitem',
            foreground=COLORS['green'],
            lmargin1=24, lmargin2=24)
        self.text.tag_config('image',
            foreground=COLORS['yellow'],
            font=('Arial', 10, 'italic'))
        self.text.tag_config('selected',
            background=COLORS['border'],
            foreground='#ffffff')

        # Selection state
        self.sel_start    = None
        self.sel_end      = None
        self.click_count  = 0
        self.line_map     = []

        # Populate
        self._populate_reader()

        # Click to select — ONLY inside this text widget
        self.text.bind('<Button-1>', self._on_text_click)
        self.text.config(state='disabled')

    def _populate_reader(self):
        self.text.config(state='normal')
        self.text.delete('1.0', 'end')
        self.line_map = []

        if not self.captured_blocks:
            self.text.insert('end', 'No content was captured.')
            self.text.config(state='disabled')
            return

        for i, block in enumerate(self.captured_blocks):
            btype = block.get('type', 'text')
            text  = block.get('text', '')
            href  = block.get('href', '')
            alt   = block.get('alt', '')

            start = self.text.index('end')

            if btype == 'heading':
                self.text.insert('end', text + '\n', 'heading')
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

    def _on_text_click(self, event):
        """Handle click inside text area ONLY — never touches other windows."""
        idx = self.text.index(f'@{event.x},{event.y}')

        if self.click_count == 0:
            self.sel_start  = idx
            self.sel_end    = None
            self.click_count = 1
            self._highlight()
            line = idx.split('.')[0]
            self.status_var.set(
                f"START set at line {line}  —  Now scroll and click where you want to STOP.")
        else:
            self.sel_end    = idx
            self.click_count = 0
            self._highlight()
            l1 = int(self.sel_start.split('.')[0])
            l2 = int(self.sel_end.split('.')[0])
            if l2 < l1:
                l1, l2 = l2, l1
            count = l2 - l1
            self.status_var.set(
                f"Selected {count} lines  —  Use Copy or Save above.")

    def _highlight(self):
        self.text.config(state='normal')
        self.text.tag_remove('selected', '1.0', 'end')
        if self.sel_start and self.sel_end:
            s, e = self.sel_start, self.sel_end
            if self.text.compare(s, '>', e):
                s, e = e, s
            self.text.tag_add('selected', s, e)
        self.text.config(state='disabled')

    def _clear_selection(self):
        self.sel_start   = None
        self.sel_end     = None
        self.click_count = 0
        self.text.config(state='normal')
        self.text.tag_remove('selected', '1.0', 'end')
        self.text.config(state='disabled')
        self.status_var.set("Selection cleared. Click to set a new START point.")

    def _get_selected_blocks(self):
        if not self.sel_start or not self.sel_end:
            return self.captured_blocks
        s, e = self.sel_start, self.sel_end
        if self.text.compare(s, '>', e):
            s, e = e, s
        selected = []
        for (bstart, bend, idx) in self.line_map:
            if (self.text.compare(bstart, '<=', e) and
                self.text.compare(bend,   '>=', s)):
                selected.append(self.captured_blocks[idx])
        return selected if selected else self.captured_blocks

    def _get_text(self):
        blocks = self._get_selected_blocks()
        lines  = []
        for b in blocks:
            btype = b.get('type','text')
            if btype == 'heading':
                lines.append('\n' + b.get('text','').upper() + '\n')
            elif btype == 'listitem':
                lines.append('  • ' + b.get('text',''))
            elif btype == 'link':
                href = b.get('href','')
                lines.append(b.get('text','') + (f'  [{href}]' if href else ''))
            elif btype == 'image':
                lines.append(f"[Image: {b.get('alt','')}]")
            else:
                lines.append(b.get('text',''))
        return '\n'.join(lines)

    def _copy_selection(self):
        text = self._get_text()
        if not text.strip():
            self.status_var.set("Nothing selected to copy.")
            return
        if HAS_CLIP:
            pyperclip.copy(text)
        else:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        self.status_var.set("Copied to clipboard!")

    def _save(self, fmt):
        cfg    = load_config()
        folder = cfg.get('save_folder', SAVE_FOLDER)
        os.makedirs(folder, exist_ok=True)
        safe   = re.sub(r'[\\/:*?"<>|]', '_',
                        self.selected_title or 'capture')[:60]
        blocks = self._get_selected_blocks()

        if fmt == 'txt':
            path = os.path.join(folder, safe + '.txt')
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self._get_text())

        elif fmt == 'html':
            path = os.path.join(folder, safe + '.html')
            _save_html(path, self.selected_title, blocks)

        elif fmt == 'docx':
            if not HAS_DOCX:
                self.status_var.set(
                    "DOCX not available. Use TXT or HTML.")
                return
            path = os.path.join(folder, safe + '.docx')
            _save_docx(path, self.selected_title, blocks)

        self.status_var.set(f"Saved to: {path}")
        try:
            os.startfile(folder)
        except Exception:
            pass

    def _start_over(self):
        self.root.geometry("500x420")
        self.selected_hwnd   = None
        self.selected_title  = ''
        self.captured_blocks = []
        self._show_welcome()

    # ── UI helpers ────────────────────────────────────────────────────────────

    def _clear(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def _header(self, title, subtitle=''):
        hdr = tk.Frame(self.root, bg=COLORS['panel'], pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text=title,
                 bg=COLORS['panel'], fg=COLORS['purple'],
                 font=('Arial', 14, 'bold'),
                 padx=16).pack(anchor='w')
        if subtitle:
            tk.Label(hdr, text=subtitle,
                     bg=COLORS['panel'], fg=COLORS['muted'],
                     font=('Arial', 10),
                     padx=16).pack(anchor='w')

    def _footer(self, back_cmd, next_cmd):
        footer = tk.Frame(self.root, bg=COLORS['bg'], pady=8)
        footer.pack(fill='x', side='bottom', padx=16)

        # Always show a big visible Cancel/Exit button
        tk.Button(footer,
                  text="✕  Exit PageCapture",
                  bg=COLORS['red'], fg=COLORS['bg'],
                  font=('Arial', 10, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4,
                  command=self.root.quit).pack(side='right', padx=4)

        if back_cmd:
            tk.Button(footer,
                      text="← Back",
                      bg=COLORS['panel'], fg=COLORS['text'],
                      font=('Arial', 10),
                      relief='flat', cursor='hand2',
                      padx=10, pady=4,
                      command=back_cmd).pack(side='left', padx=4)

# ── Save helpers ──────────────────────────────────────────────────────────────

def _esc(s):
    return str(s or '').replace('&','&amp;').replace('<','&lt;') \
                       .replace('>','&gt;').replace('"','&quot;')

def _save_html(path, title, blocks):
    lines = []
    for b in blocks:
        btype = b.get('type','text')
        if btype == 'heading':
            lines.append(f'<h2>{_esc(b.get("text",""))}</h2>')
        elif btype == 'link':
            lines.append(
                f'<p><a href="{_esc(b.get("href",""))}">'
                f'{_esc(b.get("text",""))}</a></p>')
        elif btype == 'listitem':
            lines.append(f'<li>{_esc(b.get("text",""))}</li>')
        elif btype == 'image':
            lines.append(
                f'<p><em>[Image: {_esc(b.get("alt",""))}]</em></p>')
        else:
            lines.append(f'<p>{_esc(b.get("text",""))}</p>')

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>{_esc(title)}</title>
  <style>
    body {{ font-family: Georgia, serif; max-width: 820px;
            margin: 40px auto; padding: 0 20px;
            line-height: 1.75; color: #222; }}
    h2  {{ color: #333; margin-top: 1.4em; }}
    a   {{ color: #0066cc; }}
    li  {{ margin: 0.3em 0; }}
    p   {{ margin: 0.5em 0; }}
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
    p   = doc.add_paragraph()
    run = p.add_run(title or 'Capture')
    run.bold      = True
    run.font.size = Pt(18)

    for b in blocks:
        btype = b.get('type','text')
        if btype == 'heading':
            doc.add_heading(b.get('text',''), level=2)
        elif btype == 'link':
            p   = doc.add_paragraph()
            run = p.add_run(b.get('text',''))
            run.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
            run.underline = True
            if b.get('href'):
                p.add_run(f"  [{b['href']}]")
        elif btype == 'listitem':
            doc.add_paragraph(b.get('text',''), style='List Bullet')
        elif btype == 'image':
            p   = doc.add_paragraph()
            run = p.add_run(f"[Image: {b.get('alt','')}]")
            run.italic = True
        else:
            doc.add_paragraph(b.get('text',''))

    doc.save(path)

# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    PageCaptureApp()
