import ctypes
import os
import queue
import subprocess
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from Auto_Install_CD import (
    format_recent_file_entry,
    load_config_data,
    save_config_data,
    save_base_path,
    is_valid_base_path,
    scan_target_files,
    install_to_devices,
    get_connected_devices,
    get_device_labels,
    resolve_external_install_input,
)

APP_TITLE = "QA 설치기 위젯"
WINDOW_WIDTH = 840
WINDOW_HEIGHT = 380
WINDOW_MARGIN_X = 24
WINDOW_MARGIN_Y = 64
DEFAULT_WIDGET_LOCKED = True
MIN_WINDOW_WIDTH = 720
MIN_WINDOW_HEIGHT = 310
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
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.configure(bg=BG_ROOT)
        self.root.overrideredirect(True)

        self.event_queue = queue.Queue()
        self.current_base_path = ""
        self.current_config = load_config_data()
        self.current_recent_files = []
        self.current_file_lookup = {}
        self.last_selected_file = None
        self.last_selected_ext = None
        self.busy_scan = False
        self.busy_install = False
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

        self.path_text = tk.StringVar(value="-")
        self.scan_status_text = tk.StringVar(value="대기중")

        self._build_ui()
        self._bind_context_menu()
        self.root.update_idletasks()
        self._place_window()
        self._apply_desktop_style()
        self._load_initial_config()
        self._start_device_tracker()
        self._poll_events()

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

        _tb_btn("경로 변경", self.change_path)
        _tb_btn("외부 경로 입력", self.install_external_input)
        _tb_btn("새로고침", self.refresh_all)

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

        install_button = tk.Button(
            action_frame,
            text="설치 실행",
            command=self.install_selected,
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
        install_button.pack(fill="x", padx=0, pady=(0, 0))

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
            height=4,
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
        self._save_widget_settings()
        self._refresh_info_labels()
        self.refresh_all()

    def _refresh_info_labels(self):
        self.path_text.set(self.current_base_path or "-")

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

    def refresh_all(self):
        self.refresh_devices()
        self.refresh_recent_files()

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

            files = scan_target_files(self.current_base_path, progress_callback=scan_callback)
            self.event_queue.put(("scan_done", files))

        threading.Thread(target=worker, daemon=True).start()

    def change_path(self):
        self._prompt_for_path(first_time=False)

    def install_external_input(self):
        if self.busy_install:
            return
        raw_input = simpledialog.askstring(APP_TITLE, "링크(https) 또는 파일 경로(C:\\)를 입력하세요.", parent=self.root)
        if not raw_input:
            return
        self.scan_status_text.set("외부 경로 확인 중")
        self._append_log("외부 경로 확인 중", "info")

        def worker():
            sel_file, ext, error_message = resolve_external_install_input(raw_input)
            self.event_queue.put(("external_input_ready", (sel_file, ext, error_message)))

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

            result = install_to_devices(sel_file, ext, progress_callback=progress_callback)
            self.event_queue.put(("install_done", result))

        threading.Thread(target=worker, daemon=True).start()

    def _handle_scan_done(self, files):
        self.busy_scan = False
        self.current_recent_files = files
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
            self._add_build_card(entry)

        if files:
            first_key = next(iter(self.current_file_lookup.keys()))
            self._on_build_card_click(first_key)

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
                sel_file, ext, error_message = payload
                if error_message:
                    messagebox.showerror(APP_TITLE, error_message)
                    self.scan_status_text.set(error_message)
                    self._append_log(error_message, "error")
                else:
                    self._start_install(sel_file, ext)
            elif event_name == "device_progress":
                display_name, percent, message = payload
                self._append_progress_log(display_name, percent, message)
            elif event_name == "install_done":
                self.busy_install = False
                summary = payload.get("summary", "작업 완료")
                results = payload.get("results", [])
                success_count = 0
                failure_count = 0

                for display_name, status, _ in results:
                    if display_name not in self.device_display_order:
                        self.device_display_order.append(display_name)
                    if "❌" in status:
                        failure_count += 1
                        self.device_status_map[display_name] = {"percent": 100, "message": "설치 실패"}
                    else:
                        success_count += 1
                        self.device_status_map[display_name] = {"percent": 100, "message": "설치 완료"}

                self._render_device_list()

                if payload.get("mode") == "bat":
                    final_log = summary
                    log_tag = "success"
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
