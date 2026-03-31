import ctypes
import os
import queue
import subprocess
import sys
import threading
import time


def _configure_tcl_environment():
    base_prefix = getattr(sys, "base_prefix", sys.prefix)
    tcl_root = os.path.join(base_prefix, "tcl")
    tcl_library = os.path.join(tcl_root, "tcl8.6")
    tk_library = os.path.join(tcl_root, "tk8.6")

    # Some Python 3.13 Windows installs do not resolve Tcl/Tk paths automatically.
    if not os.environ.get("TCL_LIBRARY") and os.path.exists(os.path.join(tcl_library, "init.tcl")):
        os.environ["TCL_LIBRARY"] = tcl_library
    if not os.environ.get("TK_LIBRARY") and os.path.isdir(tk_library):
        os.environ["TK_LIBRARY"] = tk_library


_configure_tcl_environment()

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

from Auto_Install_CD import (
    SUBPROCESS_NO_WINDOW,
    format_recent_file_entry,
    get_scan_parent_folders,
    get_selected_scan_folders,
    load_config_data,
    list_scan_folder_candidates,
    normalize_input_path,
    save_config_data,
    save_base_path,
    save_selected_scan_folders,
    is_valid_base_path,
    scan_target_files,
    install_to_devices,
    get_connected_devices,
    get_device_labels,
    resolve_external_install_input,
)

APP_TITLE = "QA 설치기 위젯"
WINDOW_WIDTH = 840
WINDOW_HEIGHT = 460
WINDOW_MARGIN_X = 24
WINDOW_MARGIN_Y = 64
DEFAULT_WIDGET_LOCKED = True
MIN_WINDOW_WIDTH = 720
MIN_WINDOW_HEIGHT = 390
DEFAULT_RIGHT_PANEL_WIDTH = 350

GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
HWND_BOTTOM = 1

# 디자인 토큰
BG_ROOT = "#0f1117"
BG_TITLEBAR = "#161a27"
BG_CARD = "#161a27"
BG_ITEM = "#1a1e2e"
BG_ITEM_SEL = "#1a2d3f"
BG_LOG = "#0a0d14"
BG_INPUT = "#0f1117"
BG_DEVICE = "#12151f"

FG_PRIMARY = "#e0e4f0"
FG_SECONDARY = "#8892aa"
FG_MUTED = "#4a5270"
FG_ACCENT = "#5DCAA5"
FG_SUCCESS = "#1D9E75"
FG_ERROR = "#E24B4A"
FG_WARN = "#EF9F27"
FG_COMPLETE = "#60A5FA"

BORDER_DARK = "#1e2335"
BORDER_MID = "#2a3050"
BORDER_SEL = "#1D9E75"

BTN_BG = "#1e2335"
BTN_FG = "#9aa0b8"
BTN_ACTIVE = "#252b40"

FONT_TITLE = ("Malgun Gothic", 13, "bold")
FONT_SECTION = ("Malgun Gothic", 10, "bold")
FONT_BODY = ("Malgun Gothic", 10)
FONT_SMALL = ("Malgun Gothic", 9)
FONT_MONO = ("Consolas", 10)


class DesktopInstallerWidget:
    def __init__(self):
        self.root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.configure(bg=BG_ROOT)
        self.root.overrideredirect(True)
        self.style = ttk.Style(self.root)
        self._configure_ttk_styles()

        self.event_queue = queue.Queue()
        self.current_base_path = ""
        self.current_config = load_config_data()
        self.current_recent_files = []
        self.manual_recent_files = []
        self.external_mode_active = False
        self.current_file_lookup = {}
        self.last_selected_file = None
        self.last_selected_ext = None
        self.busy_scan = False
        self.busy_install = False
        self.install_cancel_event = None
        self.device_progress_cache = {}
        self.device_status_map = {}
        self.device_display_order = []
        self.device_refresh_inflight = False
        self.device_tracker_process = None
        self.device_tracker_running = False
        self.drag_offset = (0, 0)
        self.position_locked = self._get_widget_settings().get("locked", DEFAULT_WIDGET_LOCKED)
        self.resize_origin = None
        self.selected_build_key = None
        self.scan_folder_candidate_cache_key = None
        self.scan_folder_candidate_cache = []

        self.path_text = tk.StringVar(value="-")
        self.external_input_text = tk.StringVar(value="")
        self.scan_status_text = tk.StringVar(value="대기중")

        self._build_ui()
        self._bind_context_menu()
        self.root.update_idletasks()
        self._place_window()
        self._apply_desktop_style()
        self._load_initial_config()
        self._start_device_tracker()
        self._poll_events()

    def _configure_ttk_styles(self):
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.style.configure(
            "Build.Vertical.TScrollbar",
            background=FG_SUCCESS,
            troughcolor="#000000",
            bordercolor="#000000",
            darkcolor=FG_SUCCESS,
            lightcolor=FG_SUCCESS,
            arrowcolor=FG_SUCCESS,
            gripcount=0,
            relief="flat",
            borderwidth=0,
            arrowsize=12,
        )
        self.style.map(
            "Build.Vertical.TScrollbar",
            background=[("active", FG_ACCENT), ("pressed", FG_ACCENT)],
            darkcolor=[("active", FG_ACCENT), ("pressed", FG_ACCENT)],
            lightcolor=[("active", FG_ACCENT), ("pressed", FG_ACCENT)],
            arrowcolor=[("active", FG_ACCENT), ("pressed", FG_ACCENT)],
        )
        self.style.configure(
            "Scan.Treeview",
            background=BG_INPUT,
            fieldbackground=BG_INPUT,
            foreground=FG_PRIMARY,
            borderwidth=0,
            rowheight=24,
            font=FONT_SMALL,
        )
        self.style.map(
            "Scan.Treeview",
            background=[("selected", BG_ITEM_SEL)],
            foreground=[("selected", FG_PRIMARY)],
        )

    # UI 빌드
    def _build_ui(self):
        title_bar = tk.Frame(self.root, bg=BG_TITLEBAR, pady=8)
        title_bar.pack(fill="x")
        title_bar.bind("<ButtonPress-1>", self._start_drag)
        title_bar.bind("<B1-Motion>", self._on_drag)

        title_label = tk.Label(
            title_bar,
            text="AUTO_INSTALL",
            bg=BG_TITLEBAR,
            fg=FG_ACCENT,
            font=FONT_TITLE,
            cursor="fleur",
        )
        title_label.pack(side="left", padx=(14, 8))
        title_label.bind("<ButtonPress-1>", self._start_drag)
        title_label.bind("<B1-Motion>", self._on_drag)

        self.path_entry = tk.Entry(
            title_bar,
            textvariable=self.path_text,
            relief="flat",
            bd=0,
            readonlybackground=BG_INPUT,
            fg=FG_MUTED,
            disabledforeground=FG_MUTED,
            insertbackground=FG_PRIMARY,
            font=FONT_SMALL,
            highlightthickness=1,
            highlightbackground=BORDER_MID,
            highlightcolor=BORDER_SEL,
        )
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=6)
        self.path_entry.config(state="readonly")

        def _tb_btn(text, cmd):
            button = tk.Button(
                title_bar,
                text=text,
                command=cmd,
                bg=BTN_BG,
                fg=BTN_FG,
                activebackground=BTN_ACTIVE,
                activeforeground=FG_PRIMARY,
                relief="flat",
                bd=0,
                padx=10,
                pady=4,
                font=FONT_SMALL,
                cursor="hand2",
            )
            button.pack(side="left", padx=(0, 4))
            return button

        _tb_btn("기준 위치 변경", self.change_path)
        _tb_btn("검색 범위 선택", self.change_scan_targets)
        _tb_btn("새로고침", self.user_refresh_all)

        tk.Button(
            title_bar,
            text="✕",
            command=self._close_widget,
            bg=FG_ERROR,
            fg="white",
            activebackground="#c0392b",
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="right", padx=(0, 10))

        tk.Frame(self.root, bg=BORDER_DARK, height=1).pack(fill="x")

        external_row = tk.Frame(self.root, bg=BG_ROOT, pady=6)
        external_row.pack(fill="x", padx=12)

        tk.Label(
            external_row,
            text="직접 설치",
            bg=BG_ROOT,
            fg=FG_MUTED,
            font=FONT_SMALL,
        ).pack(side="left", padx=(0, 8))

        self.external_input_entry = tk.Entry(
            external_row,
            textvariable=self.external_input_text,
            relief="flat",
            bd=0,
            bg=BG_INPUT,
            fg=FG_PRIMARY,
            insertbackground=FG_PRIMARY,
            font=FONT_SMALL,
            highlightthickness=1,
            highlightbackground=BORDER_MID,
            highlightcolor=BORDER_SEL,
        )
        self.external_input_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.external_input_entry.bind("<Return>", lambda _event: self.install_external_input())

        tk.Button(
            external_row,
            text="삭제",
            command=self._clear_external_inputs,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="left", padx=(8, 0))

        tk.Label(
            external_row,
            text="경로 입력 또는 파일 드롭",
            bg=BG_ROOT,
            fg=FG_MUTED,
            font=FONT_SMALL,
        ).pack(side="left", padx=(8, 0))

        self._bind_external_drop_target(external_row)
        self._bind_external_drop_target(self.external_input_entry)

        content = tk.Frame(self.root, bg=BG_ROOT, height=210)
        content.pack(side="top", fill="both", expand=True, padx=12, pady=(8, 4))
        content.pack_propagate(False)

        self.main_pane = tk.PanedWindow(
            content,
            orient=tk.HORIZONTAL,
            sashwidth=6,
            bg=BG_ROOT,
            bd=0,
            relief="flat",
        )
        self.main_pane.pack(fill="both", expand=True)

        left_card = tk.Frame(self.main_pane, bg=BG_CARD)
        left_card.configure(highlightthickness=1, highlightbackground=BORDER_DARK)
        self.main_pane.add(left_card, stretch="always", minsize=320)

        left_header = tk.Frame(left_card, bg=BG_CARD, pady=0)
        left_header.pack(fill="x", padx=0)

        tk.Label(
            left_header,
            text="최근 빌드",
            bg=BG_CARD,
            fg=FG_ACCENT,
            font=FONT_SECTION,
        ).pack(side="left")

        tk.Label(
            left_header,
            text="TOP 5",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_BODY,
            padx=0,
            pady=0,
        ).pack(side="left", padx=(6, 0))

        tk.Frame(left_card, bg=BORDER_DARK, height=1).pack(fill="x")

        scroll_container = tk.Frame(left_card, bg=BG_CARD)
        scroll_container.pack(fill="both", expand=True)

        self.build_canvas = tk.Canvas(
            scroll_container,
            bg=BG_CARD,
            highlightthickness=0,
            bd=0,
        )
        self.build_canvas.pack(side="left", fill="both", expand=True)

        self.build_scrollbar = ttk.Scrollbar(
            scroll_container,
            orient="vertical",
            command=self.build_canvas.yview,
            style="Build.Vertical.TScrollbar",
        )
        self.build_scrollbar.pack(side="right", fill="y")
        self.build_canvas.configure(yscrollcommand=self.build_scrollbar.set)

        self.build_list_frame = tk.Frame(self.build_canvas, bg=BG_CARD)
        self.build_canvas_window = self.build_canvas.create_window(
            (0, 0),
            window=self.build_list_frame,
            anchor="nw",
        )
        self.build_list_frame.bind(
            "<Configure>",
            lambda _e: self.build_canvas.configure(scrollregion=self.build_canvas.bbox("all")),
        )
        self.build_canvas.bind(
            "<Configure>",
            lambda event: self.build_canvas.itemconfigure(self.build_canvas_window, width=event.width),
        )
        self._bind_build_mousewheel(scroll_container)

        self.right_panel = tk.Frame(
            self.main_pane,
            bg=BG_CARD,
            width=DEFAULT_RIGHT_PANEL_WIDTH,
        )
        self.right_panel.configure(highlightthickness=1, highlightbackground=BORDER_DARK)
        self.main_pane.add(self.right_panel, minsize=300)
        self.main_pane.bind(
            "<ButtonRelease-1>",
            lambda _e: self._save_widget_settings() if not self.position_locked else None,
        )

        right_header = tk.Frame(self.right_panel, bg=BG_CARD, pady=0)
        right_header.pack(fill="x", padx=0)

        tk.Label(
            right_header,
            text="연결 디바이스",
            bg=BG_CARD,
            fg=FG_ACCENT,
            font=FONT_SECTION,
        ).pack(side="left")

        tk.Frame(self.right_panel, bg=BORDER_DARK, height=1).pack(fill="x")

        action_frame = tk.Frame(self.right_panel, bg=BG_CARD)
        action_frame.pack(side="bottom", fill="x")

        tk.Frame(action_frame, bg=BORDER_DARK, height=1).pack(fill="x")

        self.install_button = tk.Button(
            action_frame,
            text="설치 실행",
            command=self._handle_install_button,
            bg=FG_SUCCESS,
            fg="#E1F5EE",
            activebackground="#0F6E56",
            activeforeground="#E1F5EE",
            relief="flat",
            bd=0,
            font=("Malgun Gothic", 10, "bold"),
            cursor="hand2",
            pady=8,
        )
        self.install_button.pack(fill="x", padx=0, pady=(0, 0))
        self._update_install_button()

        self.status_label = tk.Label(
            action_frame,
            textvariable=self.scan_status_text,
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_BODY,
        )
        self.status_label.pack(anchor="center", pady=(0, 0))

        device_wrapper = tk.Frame(self.right_panel, bg=BG_DEVICE, pady=0, padx=0)
        device_wrapper.pack(fill="both", expand=True, padx=0, pady=(0, 0))

        self.device_listbox = tk.Listbox(
            device_wrapper,
            bg=BG_DEVICE,
            fg=FG_PRIMARY,
            selectbackground=BG_ITEM_SEL,
            selectforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            activestyle="none",
            font=FONT_BODY,
            highlightthickness=0,
        )
        self.device_listbox.pack(fill="both", expand=True)

        tk.Frame(self.root, bg=BORDER_DARK, height=1).pack(side="bottom", fill="x")

        log_section = tk.Frame(self.root, bg=BG_ROOT)
        log_section.pack(side="bottom", fill="x", expand=False, padx=12, pady=(4, 10))

        log_header = tk.Frame(log_section, bg=BG_ROOT)
        log_header.pack(fill="x", pady=(0, 6))
        tk.Label(
            log_header,
            text="로그",
            bg=BG_ROOT,
            fg=FG_ACCENT,
            font=FONT_SECTION,
        ).pack(side="left")

        self.log_text = tk.Text(
            log_section,
            height=8,
            bg=BG_LOG,
            fg=FG_SECONDARY,
            relief="flat",
            bd=0,
            wrap="word",
            font=FONT_MONO,
            insertbackground=FG_PRIMARY,
            highlightthickness=1,
            highlightbackground=BORDER_DARK,
        )
        self.log_text.pack(fill="x", expand=False)
        self.log_text.tag_configure("error", foreground=FG_ERROR)
        self.log_text.tag_configure("success", foreground=FG_SUCCESS)
        self.log_text.tag_configure("info", foreground=FG_SECONDARY)
        self.log_text.config(state="disabled")

        resize_grip = tk.Label(
            self.root,
            text="◢",
            bg=BG_ROOT,
            fg=BORDER_MID,
            cursor="size_nw_se",
            font=FONT_SMALL,
        )
        resize_grip.place(relx=1.0, rely=1.0, x=-8, y=-8, anchor="se")
        resize_grip.bind("<ButtonPress-1>", self._start_resize)
        resize_grip.bind("<B1-Motion>", self._on_resize)

    def _add_build_card(self, entry):
        display_key = entry["display_key"]

        card = tk.Frame(
            self.build_list_frame,
            bg=BG_ITEM,
            padx=0,
            pady=0,
            cursor="hand2",
        )
        card._build_key = display_key
        card.pack(fill="x", padx=0, pady=0)

        accent_bar = tk.Frame(card, bg=BORDER_MID, width=3)
        accent_bar._accent = True
        accent_bar.pack(side="left", fill="y", padx=(0, 0))

        text_frame = tk.Frame(card, bg=BG_ITEM)
        text_frame.pack(side="left", fill="x", expand=True)

        top_line = tk.Label(
            text_frame,
            text=entry["directory"],
            anchor="w",
            bg=BG_ITEM,
            fg=FG_MUTED,
            font=FONT_BODY,
            justify="left",
        )
        top_line._muted = True
        top_line.pack(fill="x")

        bottom_text = f"{entry['filename']}  ({entry['timestamp']})"
        bottom_line = tk.Label(
            text_frame,
            text=bottom_text,
            anchor="w",
            bg=BG_ITEM,
            fg=FG_PRIMARY,
            font=FONT_BODY,
            justify="left",
        )
        bottom_line.pack(fill="x", pady=(0, 0))

        def _update_wrap(event):
            wrap_width = max(140, event.width - 4)
            top_line.configure(wraplength=wrap_width)
            bottom_line.configure(wraplength=wrap_width)

        text_frame.bind("<Configure>", _update_wrap)

        for widget in (card, accent_bar, text_frame, top_line, bottom_line):
            widget.bind("<Button-1>", lambda _e, key=display_key: self._on_build_card_click(key))
            widget.bind("<MouseWheel>", self._on_build_mousewheel)
            widget.bind("<Button-4>", self._on_build_mousewheel)
            widget.bind("<Button-5>", self._on_build_mousewheel)

    def _bind_build_mousewheel(self, widget):
        for event_name in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(event_name, self._on_build_mousewheel)
            self.build_canvas.bind(event_name, self._on_build_mousewheel)
            self.build_list_frame.bind(event_name, self._on_build_mousewheel)

    def _bind_external_drop_target(self, widget):
        if not DND_FILES or not hasattr(widget, "drop_target_register"):
            return
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", self._handle_external_drop)

    def _extract_drop_paths(self, drop_data):
        if not drop_data:
            return []
        try:
            items = list(self.root.tk.splitlist(drop_data))
        except tk.TclError:
            items = [drop_data]
        paths = []
        for item in items:
            normalized_path = normalize_input_path(item)
            if normalized_path:
                paths.append(normalized_path)
        return paths

    def _handle_external_drop(self, event):
        drop_paths = self._extract_drop_paths(getattr(event, "data", ""))
        if not drop_paths:
            self.scan_status_text.set("드롭된 파일 경로를 읽지 못했습니다.")
            self._append_log("드롭된 파일 경로를 읽지 못했습니다.", "error")
            return "break"
        self.external_input_text.set(drop_paths[0])
        self._queue_external_inputs(drop_paths)
        return "break"

    def _on_build_mousewheel(self, event):
        if getattr(event, "delta", 0):
            step = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            step = -1
        elif getattr(event, "num", None) == 5:
            step = 1
        else:
            step = 0

        if step:
            self.build_canvas.yview_scroll(step, "units")
        return "break"

    def _select_build_card(self, display_key):
        self.selected_build_key = display_key
        for child in self.build_list_frame.winfo_children():
            selected = getattr(child, "_build_key", None) == display_key
            card_bg = BG_ITEM_SEL if selected else BG_ITEM
            child.configure(bg=card_bg)
            for grandchild in child.winfo_children():
                if getattr(grandchild, "_accent", False):
                    grandchild.configure(bg=BORDER_SEL if selected else BORDER_MID)
                    continue
                grandchild.configure(bg=card_bg)
                for gg in grandchild.winfo_children():
                    if isinstance(gg, tk.Label):
                        is_muted = getattr(gg, "_muted", False)
                        if selected:
                            gg.configure(bg=card_bg, fg=FG_ACCENT if is_muted else FG_PRIMARY)
                        else:
                            gg.configure(bg=card_bg, fg=FG_MUTED if is_muted else FG_PRIMARY)

    def _clear_build_cards(self):
        for widget in self.build_list_frame.winfo_children():
            widget.destroy()

    def _place_window(self):
        widget_settings = self._get_widget_settings()
        saved_x = widget_settings.get("x")
        saved_y = widget_settings.get("y")
        saved_width = widget_settings.get("width")
        saved_height = widget_settings.get("height")
        saved_right_panel_width = min(widget_settings.get("right_panel_width", DEFAULT_RIGHT_PANEL_WIDTH), DEFAULT_RIGHT_PANEL_WIDTH)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        if isinstance(saved_x, int) and isinstance(saved_y, int):
            target_width = min(WINDOW_WIDTH, max(MIN_WINDOW_WIDTH, saved_width)) if isinstance(saved_width, int) else WINDOW_WIDTH
            target_height = min(WINDOW_HEIGHT, max(MIN_WINDOW_HEIGHT, saved_height)) if isinstance(saved_height, int) else WINDOW_HEIGHT
            x = max(0, min(saved_x, screen_width - target_width))
            y = max(0, min(saved_y, screen_height - target_height))
        else:
            target_width = WINDOW_WIDTH
            target_height = WINDOW_HEIGHT
            x = screen_width - WINDOW_WIDTH - WINDOW_MARGIN_X
            y = screen_height - WINDOW_HEIGHT - WINDOW_MARGIN_Y
        self.root.geometry(f"{target_width}x{target_height}+{max(x, 0)}+{max(y, 0)}")
        self.root.update_idletasks()
        total_width = max(self.root.winfo_width(), target_width)
        left_width = max(320, total_width - int(saved_right_panel_width) - 24)
        self.main_pane.sash_place(0, left_width, 0)

    def _apply_desktop_style(self):
        hwnd = self.root.winfo_id()
        user32 = ctypes.windll.user32
        ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        ex_style = (ex_style | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style)
        self._send_to_bottom()
        self.root.after(2000, self._keep_bottom)

    def _send_to_bottom(self):
        hwnd = self.root.winfo_id()
        ctypes.windll.user32.SetWindowPos(
            hwnd,
            HWND_BOTTOM,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
        )

    def _keep_bottom(self):
        if self.root.winfo_exists():
            self._send_to_bottom()
            self.root.after(2000, self._keep_bottom)

    def _bind_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="", command=self._toggle_lock)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="새로고침", command=self.refresh_all)
        self.context_menu.add_command(label="닫기", command=self._close_widget)
        self._update_context_menu_label()
        self.root.bind_all("<Button-3>", self._show_context_menu)

    def _show_context_menu(self, event):
        self._update_context_menu_label()
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _update_context_menu_label(self):
        label = "위치 고정 풀기" if self.position_locked else "위치 고정"
        self.context_menu.entryconfig(0, label=label)

    def _get_widget_settings(self):
        settings = self.current_config.get("widget_settings", {})
        if not isinstance(settings, dict):
            settings = {}
        return settings

    def _save_widget_settings(self):
        widget_settings = self._get_widget_settings()
        widget_settings["locked"] = self.position_locked
        widget_settings["x"] = self.root.winfo_x()
        widget_settings["y"] = self.root.winfo_y()
        widget_settings["width"] = self.root.winfo_width()
        widget_settings["height"] = self.root.winfo_height()
        widget_settings["right_panel_width"] = self.right_panel.winfo_width()
        self.current_config["widget_settings"] = widget_settings
        save_config_data(self.current_config)

    def _toggle_lock(self):
        self.position_locked = not self.position_locked
        self._save_widget_settings()
        self._update_context_menu_label()
        self._send_to_bottom()
        self._append_log("위치 고정 적용" if self.position_locked else "위치 고정 해제", "info")

    def _start_drag(self, event):
        if self.position_locked:
            return
        self.drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

    def _on_drag(self, event):
        if self.position_locked:
            return
        offset_x, offset_y = self.drag_offset
        self.root.geometry(f"+{event.x_root - offset_x}+{event.y_root - offset_y}")
        self._send_to_bottom()

    def _finalize_drag_position(self):
        if not self.position_locked:
            self._save_widget_settings()

    def _start_resize(self, event):
        if self.position_locked:
            return
        self.resize_origin = (
            event.x_root,
            event.y_root,
            self.root.winfo_width(),
            self.root.winfo_height(),
        )

    def _on_resize(self, event):
        if self.position_locked or not self.resize_origin:
            return
        start_x, start_y, start_w, start_h = self.resize_origin
        delta_x = event.x_root - start_x
        delta_y = event.y_root - start_y
        new_width = max(MIN_WINDOW_WIDTH, start_w + delta_x)
        new_height = max(MIN_WINDOW_HEIGHT, start_h + delta_y)
        self.root.geometry(f"{new_width}x{new_height}")
        self._send_to_bottom()
        self._save_widget_settings()

    def _load_initial_config(self):
        config_data = self.current_config
        base_path = config_data.get("onedrive_path", "")
        if not is_valid_base_path(base_path):
            self._prompt_for_path(first_time=True)
            return
        self.current_base_path = base_path
        self.current_config = config_data
        self._refresh_info_labels()
        self.refresh_all()

    def _prompt_for_path(self, first_time=False):
        message = "최상위 폴더 경로를 선택하세요."
        if first_time:
            messagebox.showinfo(APP_TITLE, message)
        selected_dir = filedialog.askdirectory(title="최상위 폴더 경로 선택")
        if not selected_dir:
            if first_time:
                self.root.after(100, self.root.destroy)
            return
        normalized_path, config_data = save_base_path(selected_dir)
        self.current_base_path = normalized_path
        self.current_config = config_data
        self._invalidate_scan_folder_candidate_cache()
        self._save_widget_settings()
        self._refresh_info_labels()
        self.refresh_all()

    def _refresh_info_labels(self):
        self.path_text.set(self.current_base_path or "-")

    def _invalidate_scan_folder_candidate_cache(self):
        self.scan_folder_candidate_cache_key = None
        self.scan_folder_candidate_cache = []

    def _get_scan_folder_candidates(self, force_refresh=False):
        cache_key = (
            normalize_input_path(self.current_base_path),
            tuple(get_scan_parent_folders(self.current_config)),
        )
        if force_refresh or self.scan_folder_candidate_cache_key != cache_key:
            self.scan_folder_candidate_cache = list_scan_folder_candidates(
                self.current_base_path,
                self.current_config,
            )
            self.scan_folder_candidate_cache_key = cache_key
        return list(self.scan_folder_candidate_cache)

    def _normalize_scan_selection(self, selected_items):
        normalized_paths = []
        selected_lookup = set()

        for path in sorted(
            {str(item).replace("/", "\\").strip("\\") for item in selected_items if str(item).strip()},
            key=lambda value: (value.count("\\"), value.lower()),
        ):
            ancestors = []
            current = path
            while "\\" in current:
                current = current.rsplit("\\", 1)[0]
                ancestors.append(current)

            if any(ancestor in selected_lookup for ancestor in ancestors):
                continue

            normalized_paths.append(path)
            selected_lookup.add(path)

        return normalized_paths

    def change_scan_targets(self):
        if not is_valid_base_path(self.current_base_path):
            self._prompt_for_path(first_time=False)
            if not is_valid_base_path(self.current_base_path):
                return

        candidates = self._get_scan_folder_candidates()
        if not candidates:
            parent_folders = ", ".join(get_scan_parent_folders(self.current_config))
            messagebox.showinfo(
                APP_TITLE,
                f"선택 가능한 스캔 하위 폴더가 없습니다.\n기준 폴더: {parent_folders}",
            )
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("검색 범위 선택")
        dialog.configure(bg=BG_CARD)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        dialog.geometry("760x620")

        tk.Label(
            dialog,
            text="기준 폴더와 그 하위 폴더 중 검색할 위치를 선택하세요.",
            bg=BG_CARD,
            fg=FG_PRIMARY,
            font=FONT_SECTION,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(14, 6))

        tk.Label(
            dialog,
            text="상위 폴더를 선택하면 그 아래 전체를 검색합니다.",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 10))

        tk.Label(
            dialog,
            text="좌측 트리에서 폴더를 더블클릭하면 선택 요약에 추가되거나 제거됩니다.",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 10))

        tk.Label(
            dialog,
            text="선택을 비우고 저장하면 전체 검색으로 동작합니다.",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 10))

        search_var = tk.StringVar(value="")
        summary_var = tk.StringVar(value="")
        candidate_set = set(candidates)
        saved_paths = {
            path for path in get_selected_scan_folders(self.current_config)
            if path in candidate_set
        }
        draft_selected_paths = set()
        reset_requested = False
        tree_item_by_path = {}
        all_paths = sorted(candidates, key=lambda value: (value.count("\\"), value.lower()))

        search_row = tk.Frame(dialog, bg=BG_CARD)
        search_row.pack(fill="x", padx=14, pady=(0, 10))

        tk.Label(
            search_row,
            text="검색",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
        ).pack(side="left", padx=(0, 8))

        search_entry = tk.Entry(
            search_row,
            textvariable=search_var,
            relief="flat",
            bd=0,
            bg=BG_INPUT,
            fg=FG_PRIMARY,
            insertbackground=FG_PRIMARY,
            font=FONT_SMALL,
            highlightthickness=1,
            highlightbackground=BORDER_MID,
            highlightcolor=BORDER_SEL,
        )
        search_entry.pack(side="left", fill="x", expand=True, ipady=6)

        tk.Label(
            dialog,
            text=f"현재 적용 중인 저장값: {len(self._normalize_scan_selection(saved_paths))}개",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 6))

        tk.Label(
            dialog,
            text="표시: [x] 이번 창에서 직접 선택 / [>] 이번 창 선택으로 포함 / [*] 저장된 기본값 / [~] 저장된 기본값으로 포함 / [ ] 미선택",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 10))

        content_frame = tk.Frame(dialog, bg=BG_CARD)
        content_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        tree_frame = tk.Frame(content_frame, bg=BG_CARD)
        tree_frame.pack(side="left", fill="both", expand=True)

        tree = ttk.Treeview(
            tree_frame,
            show="tree",
            style="Scan.Treeview",
            selectmode="browse",
        )
        tree_scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=tree.yview,
            style="Build.Vertical.TScrollbar",
        )
        tree.configure(yscrollcommand=tree_scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        tree_scrollbar.pack(side="right", fill="y")

        summary_frame = tk.Frame(content_frame, bg=BG_CARD, width=240)
        summary_frame.pack(side="left", fill="both", padx=(12, 0))
        summary_frame.pack_propagate(False)

        tk.Label(
            summary_frame,
            text="선택 요약",
            bg=BG_CARD,
            fg=FG_PRIMARY,
            font=FONT_SECTION,
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            summary_frame,
            textvariable=summary_var,
            bg=BG_CARD,
            fg=FG_MUTED,
            font=FONT_SMALL,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(4, 8))

        summary_listbox = tk.Listbox(
            summary_frame,
            bg=BG_INPUT,
            fg=FG_PRIMARY,
            relief="flat",
            bd=0,
            font=FONT_SMALL,
            highlightthickness=1,
            highlightbackground=BORDER_MID,
            highlightcolor=BORDER_SEL,
        )
        summary_scrollbar = ttk.Scrollbar(
            summary_frame,
            orient="vertical",
            command=summary_listbox.yview,
            style="Build.Vertical.TScrollbar",
        )
        summary_listbox.configure(yscrollcommand=summary_scrollbar.set)
        summary_listbox.pack(side="left", fill="both", expand=True)
        summary_scrollbar.pack(side="right", fill="y")

        def get_parent_path(path):
            return path.rsplit("\\", 1)[0] if "\\" in path else ""

        def get_display_name(path):
            return path.rsplit("\\", 1)[-1]

        def find_ancestor_in_selection(path, selection_set):
            current = get_parent_path(path)
            while current:
                if current in selection_set:
                    return current
                current = get_parent_path(current)
            return ""

        def get_selection_state(path):
            if path in draft_selected_paths:
                return "draft_explicit"
            if find_ancestor_in_selection(path, draft_selected_paths):
                return "draft_included"
            if path in saved_paths:
                return "saved_explicit"
            if find_ancestor_in_selection(path, saved_paths):
                return "saved_included"
            return "none"

        def get_visible_paths(query_text):
            query = query_text.strip().lower()
            if not query:
                return set(all_paths)

            visible_paths = set()
            for path in all_paths:
                if query in path.lower():
                    current = path
                    while current:
                        visible_paths.add(current)
                        current = get_parent_path(current)
            return visible_paths

        def build_summary_labels(paths):
            name_counts = {}
            for path in paths:
                folder_name = get_display_name(path)
                name_counts[folder_name] = name_counts.get(folder_name, 0) + 1

            labels = []
            for path in paths:
                folder_name = get_display_name(path)
                if name_counts.get(folder_name, 0) == 1:
                    labels.append(folder_name)
                    continue

                parent_name = get_display_name(get_parent_path(path))
                if parent_name:
                    labels.append(f"{folder_name} ({parent_name})")
                else:
                    labels.append(folder_name)

            return labels

        def update_summary():
            summary_listbox.delete(0, "end")
            normalized_selection = self._normalize_scan_selection(draft_selected_paths)
            if normalized_selection:
                summary_var.set(f"이번 창에서 선택한 폴더 {len(normalized_selection)}개")
                for label in build_summary_labels(normalized_selection):
                    summary_listbox.insert("end", label)
            elif reset_requested:
                summary_var.set("초기화 예정: 저장 시 전체 검색")
            else:
                summary_var.set("이번 창에서 선택한 폴더 0개")

        def rebuild_tree(*_args):
            expanded_paths = set()
            for path, item_id in list(tree_item_by_path.items()):
                if tree.exists(item_id) and tree.item(item_id, "open"):
                    expanded_paths.add(path)

            tree_item_by_path.clear()
            tree.delete(*tree.get_children())

            visible_paths = get_visible_paths(search_var.get())
            show_all_expanded = bool(search_var.get().strip())

            for path in all_paths:
                if path not in visible_paths:
                    continue

                parent_path = get_parent_path(path)
                parent_item_id = tree_item_by_path.get(parent_path, "")
                selection_state = get_selection_state(path)
                marker = {
                    "draft_explicit": "[x]",
                    "draft_included": "[>]",
                    "saved_explicit": "[*]",
                    "saved_included": "[~]",
                    "none": "[ ]",
                }[selection_state]
                item_id = tree.insert(
                    parent_item_id,
                    "end",
                    text=f"{marker} {get_display_name(path)}",
                    open=show_all_expanded or path in expanded_paths,
                    values=(path,),
                )
                tree_item_by_path[path] = item_id

        def toggle_selected_path(path):
            nonlocal reset_requested
            if not path:
                return
            reset_requested = False
            if path in draft_selected_paths:
                draft_selected_paths.remove(path)
            else:
                selected_ancestor = find_ancestor_in_selection(path, draft_selected_paths)
                if selected_ancestor:
                    draft_selected_paths.remove(selected_ancestor)
                draft_selected_paths.add(path)
            rebuild_tree()
            update_summary()
            item_id = tree_item_by_path.get(path)
            if item_id:
                tree.selection_set(item_id)
                tree.focus(item_id)

        def handle_tree_double_click(event):
            item_id = tree.identify_row(event.y)
            if not item_id:
                return

            element = tree.identify_element(event.x, event.y)
            if element == "Treeitem.indicator":
                return

            path = tree.item(item_id, "values")
            if not path:
                return
            toggle_selected_path(path[0])

        def toggle_focused_item(_event=None):
            focused_item = tree.focus()
            if not focused_item:
                return
            path = tree.item(focused_item, "values")
            if not path:
                return
            toggle_selected_path(path[0])

        def clear_search():
            search_var.set("")
            search_entry.focus_set()

        def refresh_candidates():
            nonlocal all_paths
            candidates = self._get_scan_folder_candidates(force_refresh=True)
            valid_candidate_set = set(candidates)
            saved_paths.intersection_update(valid_candidate_set)
            draft_selected_paths.intersection_update(valid_candidate_set)
            all_paths = sorted(candidates, key=lambda value: (value.count("\\"), value.lower()))
            rebuild_tree()
            update_summary()

        def select_all():
            nonlocal reset_requested
            reset_requested = False
            draft_selected_paths.clear()
            draft_selected_paths.update(all_paths)
            rebuild_tree()
            update_summary()

        def reset_selection():
            nonlocal reset_requested
            reset_requested = True
            draft_selected_paths.clear()
            rebuild_tree()
            update_summary()

        def remove_summary_selection(_event=None):
            selected_indexes = list(summary_listbox.curselection())
            if not selected_indexes:
                return

            normalized_selection = self._normalize_scan_selection(draft_selected_paths)
            for index in reversed(selected_indexes):
                if 0 <= index < len(normalized_selection):
                    draft_selected_paths.discard(normalized_selection[index])

            rebuild_tree()
            update_summary()

        def save_selection():
            selected_items = self._normalize_scan_selection(draft_selected_paths)
            if reset_requested:
                self.current_config = save_selected_scan_folders([], self.current_config)
                message = "검색 폴더 초기화 완료 (기본값: 전체 검색)"
            elif selected_items:
                self.current_config = save_selected_scan_folders(selected_items, self.current_config)
                message = f"검색 폴더 {len(selected_items)}개 저장"
            else:
                message = f"변경사항 없음 (기존 검색 폴더 {len(self._normalize_scan_selection(saved_paths))}개 유지)"
            self.scan_status_text.set(message)
            self._append_log(message, "info")
            dialog.destroy()
            self.refresh_recent_files()

        search_var.trace_add("write", rebuild_tree)
        tree.bind("<Double-1>", handle_tree_double_click)
        tree.bind("<Return>", toggle_focused_item)
        tree.bind("<space>", toggle_focused_item)
        summary_listbox.bind("<Double-1>", remove_summary_selection)
        summary_listbox.bind("<Delete>", remove_summary_selection)
        rebuild_tree()
        update_summary()
        search_entry.focus_set()

        button_row = tk.Frame(dialog, bg=BG_CARD)
        button_row.pack(fill="x", padx=14, pady=(0, 14))

        tk.Button(
            search_row,
            text="검색 지우기",
            command=clear_search,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            search_row,
            text="목록 새로고침",
            command=refresh_candidates,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            button_row,
            text="전체 선택",
            command=select_all,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="left")

        tk.Button(
            button_row,
            text="초기화",
            command=reset_selection,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="left", padx=(6, 0))

        tk.Button(
            button_row,
            text="선택 제거",
            command=remove_summary_selection,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="left", padx=(6, 0))

        tk.Button(
            button_row,
            text="취소",
            command=dialog.destroy,
            bg=BTN_BG,
            fg=BTN_FG,
            activebackground=BTN_ACTIVE,
            activeforeground=FG_PRIMARY,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="right")

        tk.Button(
            button_row,
            text="저장",
            command=save_selection,
            bg=FG_SUCCESS,
            fg="white",
            activebackground="#0F6E56",
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=12,
            pady=4,
            font=FONT_SMALL,
            cursor="hand2",
        ).pack(side="right", padx=(0, 6))

    def _append_log(self, message, tag="info"):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _normalize_device_status_text(self, percent, message):
        if int(percent) >= 100:
            if "취소" in message:
                return "설치 취소"
            if "❌" in message:
                return "설치 실패"
            if "✅" in message:
                return "설치 완료"
        return message

    def _get_device_status_color(self, percent, message):
        normalized = self._normalize_device_status_text(percent, message)
        if normalized in ("설치 실패", "설치 취소"):
            return FG_ERROR
        if normalized == "설치 완료":
            return FG_COMPLETE
        if int(percent) > 0:
            return FG_WARN
        return FG_PRIMARY

    def _render_device_list(self):
        self.device_listbox.delete(0, "end")
        if not self.device_display_order:
            self.device_listbox.insert("end", "연결된 기기 없음")
            return
        for display_name in self.device_display_order:
            info = self.device_status_map.get(display_name)
            if not info:
                self.device_listbox.insert("end", display_name)
                continue
            percent = int(info.get("percent", 0))
            raw_message = info.get("message", "대기중")
            status_text = self._normalize_device_status_text(percent, raw_message)
            item_color = self._get_device_status_color(percent, raw_message)
            self.device_listbox.insert(
                "end",
                f"{display_name}  {percent:3d}% {status_text}",
            )
            item_index = self.device_listbox.size() - 1
            self.device_listbox.itemconfig(item_index, fg=item_color, selectforeground=item_color)

    def _append_progress_log(self, display_name, percent, message):
        cache_key = (int(percent), message)
        if self.device_progress_cache.get(display_name) == cache_key:
            return
        self.device_progress_cache[display_name] = cache_key
        if display_name not in self.device_display_order:
            self.device_display_order.append(display_name)
        self.device_status_map[display_name] = {
            "percent": int(percent),
            "message": message,
        }
        self._render_device_list()

    def _update_install_button(self):
        if self.busy_install:
            self.install_button.configure(
                text="설치 중지",
                bg=FG_ERROR,
                activebackground="#c0392b",
            )
        else:
            self.install_button.configure(
                text="설치 실행",
                bg=FG_SUCCESS,
                activebackground="#0F6E56",
            )

    def _handle_install_button(self):
        if self.busy_install:
            self.cancel_install()
            return
        self.install_selected()

    def cancel_install(self):
        if not self.busy_install or not self.install_cancel_event:
            return
        if self.install_cancel_event.is_set():
            return
        self.install_cancel_event.set()
        self.scan_status_text.set("설치 중지 요청 중")
        self._append_log("설치 중지 요청", "info")

    def refresh_all(self):
        self.refresh_devices()
        self.refresh_recent_files()

    def _clear_external_inputs(self):
        self.external_input_text.set("")
        self.manual_recent_files = []
        self.external_mode_active = False
        if self.busy_scan:
            return
        if self.current_recent_files:
            self._render_recent_file_cards(self.current_recent_files)
        else:
            self.current_file_lookup = {}
            self._clear_build_cards()
            self.last_selected_file = None
            self.last_selected_ext = None

    def user_refresh_all(self):
        self._clear_external_inputs()
        self._invalidate_scan_folder_candidate_cache()
        self.refresh_all()

    def refresh_devices(self):
        if self.device_refresh_inflight:
            return
        self.device_refresh_inflight = True

        def worker():
            try:
                devices = get_connected_devices()
                labels = list(get_device_labels(devices).values()) if devices else []
                self.event_queue.put(("devices_loaded", labels))
            finally:
                self.event_queue.put(("devices_refresh_finished", None))

        threading.Thread(target=worker, daemon=True).start()

    def _start_device_tracker(self):
        if self.device_tracker_running:
            return
        self.device_tracker_running = True
        threading.Thread(target=self._device_tracker_worker, daemon=True).start()

    def _stop_device_tracker(self):
        self.device_tracker_running = False
        process = self.device_tracker_process
        self.device_tracker_process = None
        if process and process.poll() is None:
            try:
                process.terminate()
                process.wait(timeout=2)
            except Exception:
                try:
                    process.kill()
                except Exception:
                    pass

    def _device_tracker_worker(self):
        while self.device_tracker_running:
            process = None
            try:
                process = subprocess.Popen(
                    ["adb", "track-devices"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1,
                    creationflags=SUBPROCESS_NO_WINDOW,
                )
                self.device_tracker_process = process
                self.event_queue.put(("device_tracker_changed", None))

                while self.device_tracker_running and process.poll() is None:
                    line = process.stdout.readline()
                    if line == "":
                        break
                    stripped = line.strip()
                    if stripped.startswith("List of devices attached"):
                        continue
                    self.event_queue.put(("device_tracker_changed", None))
            except Exception:
                pass
            finally:
                if process and process.poll() is None:
                    try:
                        process.terminate()
                    except Exception:
                        pass
                if self.device_tracker_process is process:
                    self.device_tracker_process = None

            if self.device_tracker_running:
                time.sleep(1)

    def _close_widget(self):
        self._stop_device_tracker()
        if self.root.winfo_exists():
            self.root.destroy()

    def refresh_recent_files(self):
        if self.busy_scan:
            return
        if not is_valid_base_path(self.current_base_path):
            self._prompt_for_path(first_time=False)
            return
        self.busy_scan = True
        self.scan_status_text.set("스캔 준비중")
        self._append_log("최근 빌드 목록 스캔 시작", "info")

        def worker():
            def scan_callback(info):
                self.event_queue.put(("scan_progress", info))

            files = scan_target_files(
                self.current_base_path,
                selected_folders=get_selected_scan_folders(self.current_config),
                progress_callback=scan_callback,
                config_data=self.current_config,
            )
            self.event_queue.put(("scan_done", files))

        threading.Thread(target=worker, daemon=True).start()

    def change_path(self):
        self._clear_external_inputs()
        self._prompt_for_path(first_time=False)

    def install_external_input(self):
        raw_input = self.external_input_text.get().strip()
        if not raw_input:
            messagebox.showwarning(APP_TITLE, "파일 경로 또는 링크를 입력하세요.")
            return
        self._queue_external_inputs([raw_input])

    def _queue_external_inputs(self, raw_inputs):
        normalized_inputs = [str(item).strip() for item in raw_inputs if str(item).strip()]
        if not normalized_inputs:
            return
        self.scan_status_text.set("외부 경로 확인 중")
        self._append_log("외부 경로 확인 중", "info")

        def worker():
            results = []
            for raw_input in normalized_inputs:
                sel_file, ext, error_message = resolve_external_install_input(raw_input)
                results.append((raw_input, sel_file, ext, error_message))
            self.event_queue.put(("external_input_ready", results))

        threading.Thread(target=worker, daemon=True).start()

    def install_selected(self):
        if self.busy_install:
            return
        if not self.last_selected_file or not self.last_selected_ext:
            messagebox.showwarning(APP_TITLE, "설치할 파일을 먼저 선택하세요.")
            return
        self._start_install(self.last_selected_file, self.last_selected_ext)

    def _start_install(self, sel_file, ext):
        self.busy_install = True
        self.install_cancel_event = threading.Event()
        self._update_install_button()
        self.last_selected_file = sel_file
        self.last_selected_ext = ext
        self.device_progress_cache = {}
        for display_name in self.device_display_order:
            self.device_status_map[display_name] = {
                "percent": 0,
                "message": "대기중",
            }
        self._render_device_list()
        self._append_log(f"설치 시작: {sel_file}", "info")
        self.scan_status_text.set("설치 진행중")

        def worker():
            def progress_callback(display_name, percent, message):
                self.event_queue.put(("device_progress", (display_name, percent, message)))

            result = install_to_devices(
                sel_file,
                ext,
                progress_callback=progress_callback,
                cancel_event=self.install_cancel_event,
            )
            self.event_queue.put(("install_done", result))

        threading.Thread(target=worker, daemon=True).start()

    def _upsert_manual_recent_file(self, file_path):
        normalized_path = normalize_input_path(file_path)
        if not normalized_path:
            return None
        try:
            modified_time = os.path.getmtime(normalized_path)
        except OSError:
            modified_time = time.time()
        self.manual_recent_files = [
            (existing_path, existing_time)
            for existing_path, existing_time in self.manual_recent_files
            if normalize_input_path(existing_path) != normalized_path
        ]
        self.manual_recent_files.insert(0, (normalized_path, modified_time))
        self.manual_recent_files = self.manual_recent_files[:5]
        return normalized_path

    def _render_recent_file_cards(self, files, select_path=None):
        normalized_select_path = normalize_input_path(select_path) if select_path else None
        selected_key = None
        self.current_file_lookup = {}
        self._clear_build_cards()

        for full_path, modified_time in files:
            entry = format_recent_file_entry(full_path, modified_time, self.current_base_path)
            display_key = f"{entry['directory']}|{entry['filename']}|{entry['timestamp']}"
            entry["display_key"] = display_key
            self.current_file_lookup[display_key] = {
                "path": entry["path"],
                "ext": entry["extension"],
            }
            if normalized_select_path and normalize_input_path(entry["path"]) == normalized_select_path:
                selected_key = display_key
            self._add_build_card(entry)

        if self.current_file_lookup:
            target_key = selected_key or next(iter(self.current_file_lookup.keys()))
            self._on_build_card_click(target_key)

    def _handle_scan_done(self, files):
        self.busy_scan = False
        self.current_recent_files = files
        if not self.external_mode_active:
            self._render_recent_file_cards(self.current_recent_files)

        self.scan_status_text.set(f"최근 파일 {len(files)}건 로드 완료")
        self._append_log(f"최근 파일 목록 갱신 완료 ({len(files)}건)", "success")

    def _on_build_card_click(self, display_key):
        file_info = self.current_file_lookup.get(display_key)
        if not file_info:
            return
        self.last_selected_file = file_info["path"]
        self.last_selected_ext = file_info["ext"]
        self._select_build_card(display_key)

    def _poll_events(self):
        while True:
            try:
                event_name, payload = self.event_queue.get_nowait()
            except queue.Empty:
                break

            if event_name == "scan_progress":
                self.scan_status_text.set(
                    f"{payload['phase_label']} {payload['percent']}% / "
                    f"폴더 {payload['current_count']}/{payload['total_count']} / "
                    f"파일 {payload['found_files']}개"
                )
            elif event_name == "scan_done":
                self._handle_scan_done(payload)
            elif event_name == "devices_loaded":
                previous_status_map = dict(self.device_status_map)
                self.device_display_order = list(payload)
                self.device_status_map = {}
                for label in self.device_display_order:
                    existing = previous_status_map.get(label)
                    if existing:
                        self.device_status_map[label] = existing
                    else:
                        self.device_status_map[label] = {"percent": 0, "message": "대기중"}
                self._render_device_list()
            elif event_name == "devices_refresh_finished":
                self.device_refresh_inflight = False
            elif event_name == "device_tracker_changed":
                self.refresh_devices()
            elif event_name == "external_input_ready":
                valid_files = []
                error_messages = []
                for raw_input, sel_file, _ext, error_message in payload:
                    if error_message:
                        error_messages.append(f"{raw_input} -> {error_message}")
                        continue
                    selected_path = self._upsert_manual_recent_file(sel_file)
                    if selected_path:
                        valid_files.append(selected_path)

                if valid_files:
                    self.external_mode_active = True
                    self.external_input_text.set(valid_files[0])
                    self._render_recent_file_cards(self.manual_recent_files, select_path=valid_files[0])
                    self.scan_status_text.set(f"외부 파일 {len(valid_files)}건 추가됨")
                    self._append_log(f"외부 파일 {len(valid_files)}건을 최근 빌드에 추가", "success")

                if error_messages:
                    joined_message = "\n".join(error_messages[:5])
                    if len(error_messages) > 5:
                        joined_message += "\n..."
                    if not valid_files:
                        messagebox.showerror(APP_TITLE, joined_message)
                    self.scan_status_text.set("외부 파일 처리 중 오류 발생")
                    self._append_log(joined_message, "error")
            elif event_name == "device_progress":
                display_name, percent, message = payload
                self._append_progress_log(display_name, percent, message)
            elif event_name == "install_done":
                self.busy_install = False
                self.install_cancel_event = None
                self._update_install_button()
                summary = payload.get("summary", "작업 완료")
                results = payload.get("results", [])
                success_count = 0
                failure_count = 0
                cancel_count = 0

                for display_name, status, _ in results:
                    if display_name not in self.device_display_order:
                        self.device_display_order.append(display_name)
                    if "취소" in status:
                        cancel_count += 1
                        self.device_status_map[display_name] = {"percent": 100, "message": "설치 취소"}
                    elif "❌" in status:
                        failure_count += 1
                        self.device_status_map[display_name] = {"percent": 100, "message": "설치 실패"}
                    else:
                        success_count += 1
                        self.device_status_map[display_name] = {"percent": 100, "message": "설치 완료"}

                self._render_device_list()

                if payload.get("mode") == "bat":
                    final_log = summary
                    log_tag = "success"
                elif cancel_count and success_count == 0 and failure_count == 0:
                    final_log = "설치 취소"
                    log_tag = "info"
                elif cancel_count:
                    final_log = f"{success_count}대 설치 완료 / {failure_count}대 설치 실패 / {cancel_count}대 설치 취소"
                    log_tag = "info" if failure_count == 0 else "error"
                elif failure_count == 0 and results:
                    final_log = "모든 디바이스 설치 완료"
                    log_tag = "success"
                elif failure_count > 0:
                    final_log = f"{success_count}대 설치 완료 / {failure_count}대 설치 실패"
                    log_tag = "error"
                else:
                    final_log = summary
                    log_tag = "error" if not payload.get("success", True) else "success"

                self.scan_status_text.set(final_log)
                self._append_log(final_log, log_tag)

        self.root.after(100, self._poll_events)

    def run(self):
        self.root.bind("<ButtonRelease-1>", lambda _event: self._finalize_drag_position())
        self.root.mainloop()


if __name__ == "__main__":
    DesktopInstallerWidget().run()
