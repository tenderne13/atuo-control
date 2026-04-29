import json
import sys
import threading
import time
from dataclasses import dataclass, asdict
from pathlib import Path

import ctypes
import tkinter as tk
from tkinter import messagebox, ttk


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"

SW_RESTORE = 9
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
SM_CXSCREEN = 0
SM_CYSCREEN = 1

user32 = ctypes.WinDLL("user32", use_last_error=True)


@dataclass
class ActionStep:
    name: str
    x: int
    y: int
    delay_ms: int


DEFAULT_CONFIG = {
    "window_title": "MuMu安卓设备",
    "window_width": 1920,
    "window_height": 1080,
    "window_x": 0,
    "window_y": 0,
    "loop_forever": True,
    "loop_count": 1,
    "startup_delay_ms": 0,
    "steps": [
        asdict(ActionStep("开始匹配按钮", 1822, 850, 10000)),
        asdict(ActionStep("空白区域", 1830, 720, 10000)),
        asdict(ActionStep("返回按钮", 1290, 940, 10000)),
    ],
}


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return DEFAULT_CONFIG.copy()

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except (OSError, json.JSONDecodeError):
        return DEFAULT_CONFIG.copy()

    merged = DEFAULT_CONFIG.copy()
    merged.update({k: v for k, v in data.items() if k != "steps"})
    merged["steps"] = data.get("steps", DEFAULT_CONFIG["steps"])
    return merged


def save_config(config: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as file:
        json.dump(config, file, ensure_ascii=False, indent=2)


def find_window(window_title: str) -> int:
    return user32.FindWindowW(None, window_title)


def activate_window(hwnd: int) -> None:
    user32.ShowWindow(hwnd, SW_RESTORE)
    user32.SetForegroundWindow(hwnd)


def resize_and_move_window(hwnd: int, x: int, y: int, width: int, height: int) -> None:
    user32.MoveWindow(hwnd, x, y, width, height, True)


def set_cursor_pos(x: int, y: int) -> None:
    if not user32.SetCursorPos(x, y):
        raise ctypes.WinError(ctypes.get_last_error())


def mouse_left_click() -> None:
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def click_screen(x: int, y: int) -> None:
    set_cursor_pos(x, y)
    time.sleep(0.08)
    mouse_left_click()


class AutoClickApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("模拟点击控制台")
        self.root.geometry("760x620")
        self.root.minsize(720, 560)

        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        self.config_data = load_config()

        self.window_title_var = tk.StringVar(value=self.config_data["window_title"])
        self.window_width_var = tk.StringVar(value=str(self.config_data["window_width"]))
        self.window_height_var = tk.StringVar(value=str(self.config_data["window_height"]))
        self.window_x_var = tk.StringVar(value=str(self.config_data["window_x"]))
        self.window_y_var = tk.StringVar(value=str(self.config_data["window_y"]))
        self.loop_forever_var = tk.BooleanVar(value=self.config_data["loop_forever"])
        self.loop_count_var = tk.StringVar(value=str(self.config_data["loop_count"]))
        self.startup_delay_var = tk.StringVar(value=str(self.config_data["startup_delay_ms"]))

        self._build_ui()
        self._load_steps()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        settings = ttk.LabelFrame(main, text="窗口配置", padding=12)
        settings.pack(fill="x")

        self._add_labeled_entry(settings, "窗口标题", self.window_title_var, 0, 0, width=28)
        self._add_labeled_entry(settings, "窗口宽度", self.window_width_var, 0, 2)
        self._add_labeled_entry(settings, "窗口高度", self.window_height_var, 0, 4)
        self._add_labeled_entry(settings, "窗口 X", self.window_x_var, 1, 0)
        self._add_labeled_entry(settings, "窗口 Y", self.window_y_var, 1, 2)
        self._add_labeled_entry(settings, "启动前等待(ms)", self.startup_delay_var, 1, 4)

        loop_frame = ttk.Frame(main, padding=(0, 12, 0, 12))
        loop_frame.pack(fill="x")

        ttk.Checkbutton(
            loop_frame,
            text="无限循环",
            variable=self.loop_forever_var,
            command=self._toggle_loop_count_state,
        ).pack(side="left")
        ttk.Label(loop_frame, text="循环次数").pack(side="left", padx=(16, 6))
        self.loop_count_entry = ttk.Entry(loop_frame, textvariable=self.loop_count_var, width=8)
        self.loop_count_entry.pack(side="left")

        actions = ttk.Frame(main, padding=(0, 0, 0, 10))
        actions.pack(fill="x")
        ttk.Button(actions, text="保存配置", command=self.save_current_config).pack(side="left")
        ttk.Button(actions, text="测试单次", command=self.run_once).pack(side="left", padx=8)
        ttk.Button(actions, text="开始循环", command=self.start_loop).pack(side="left")
        ttk.Button(actions, text="停止", command=self.stop_loop).pack(side="left", padx=8)

        step_frame = ttk.LabelFrame(main, text="点击步骤", padding=12)
        step_frame.pack(fill="both", expand=True)

        columns = ("name", "x", "y", "delay_ms")
        self.step_table = ttk.Treeview(step_frame, columns=columns, show="headings", height=10)
        headers = {
            "name": "步骤名",
            "x": "X",
            "y": "Y",
            "delay_ms": "点击后等待(ms)",
        }
        widths = {"name": 260, "x": 100, "y": 100, "delay_ms": 140}
        for column in columns:
            self.step_table.heading(column, text=headers[column])
            self.step_table.column(column, width=widths[column], anchor="center")

        self.step_table.pack(side="left", fill="both", expand=True)
        self.step_table.bind("<<TreeviewSelect>>", self._fill_form_from_selection)

        scroll = ttk.Scrollbar(step_frame, orient="vertical", command=self.step_table.yview)
        self.step_table.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")

        editor = ttk.LabelFrame(main, text="步骤编辑", padding=12)
        editor.pack(fill="x", pady=(12, 0))

        self.step_name_var = tk.StringVar()
        self.step_x_var = tk.StringVar()
        self.step_y_var = tk.StringVar()
        self.step_delay_var = tk.StringVar()

        self._add_labeled_entry(editor, "步骤名", self.step_name_var, 0, 0, width=24)
        self._add_labeled_entry(editor, "X", self.step_x_var, 0, 2, width=10)
        self._add_labeled_entry(editor, "Y", self.step_y_var, 0, 4, width=10)
        self._add_labeled_entry(editor, "等待(ms)", self.step_delay_var, 0, 6, width=12)

        button_row = ttk.Frame(editor)
        button_row.grid(row=1, column=0, columnspan=8, pady=(12, 0), sticky="w")
        ttk.Button(button_row, text="新增步骤", command=self.add_step).pack(side="left")
        ttk.Button(button_row, text="更新选中", command=self.update_step).pack(side="left", padx=8)
        ttk.Button(button_row, text="删除选中", command=self.delete_step).pack(side="left")
        ttk.Button(button_row, text="上移", command=lambda: self.move_step(-1)).pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="下移", command=lambda: self.move_step(1)).pack(side="left", padx=8)

        actions = ttk.Frame(main, padding=(0, 12, 0, 12))
        actions.pack_forget()
        ttk.Button(actions, text="保存配置", command=self.save_current_config).pack(side="left")
        ttk.Button(actions, text="测试单次", command=self.run_once).pack(side="left", padx=8)
        ttk.Button(actions, text="开始循环", command=self.start_loop).pack(side="left")
        ttk.Button(actions, text="停止", command=self.stop_loop).pack(side="left", padx=8)

        log_frame = ttk.LabelFrame(main, text="运行日志", padding=12)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=7, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

        self._toggle_loop_count_state()

    def _add_labeled_entry(
        self,
        parent: ttk.Widget,
        label: str,
        variable: tk.StringVar,
        row: int,
        column: int,
        width: int = 12,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=variable, width=width).grid(
            row=row, column=column + 1, sticky="w", padx=(0, 16), pady=4
        )

    def _load_steps(self) -> None:
        self.step_table.delete(*self.step_table.get_children())
        for step in self.config_data["steps"]:
            self.step_table.insert("", "end", values=(step["name"], step["x"], step["y"], step["delay_ms"]))

    def _fill_form_from_selection(self, _event=None) -> None:
        selection = self.step_table.selection()
        if not selection:
            return
        values = self.step_table.item(selection[0], "values")
        self.step_name_var.set(values[0])
        self.step_x_var.set(values[1])
        self.step_y_var.set(values[2])
        self.step_delay_var.set(values[3])

    def _toggle_loop_count_state(self) -> None:
        state = "disabled" if self.loop_forever_var.get() else "normal"
        self.loop_count_entry.configure(state=state)

    def log(self, message: str) -> None:
        def write() -> None:
            now = time.strftime("%H:%M:%S")
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"[{now}] {message}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

        self.root.after(0, write)

    def _read_int(self, value: str, field_name: str, minimum: int | None = None) -> int:
        try:
            result = int(value)
        except ValueError as exc:
            raise ValueError(f"{field_name} 必须是整数") from exc
        if minimum is not None and result < minimum:
            raise ValueError(f"{field_name} 不能小于 {minimum}")
        return result

    def collect_config(self) -> dict:
        steps = []
        for item in self.step_table.get_children():
            name, x, y, delay_ms = self.step_table.item(item, "values")
            steps.append(
                {
                    "name": str(name),
                    "x": self._read_int(str(x), "步骤 X"),
                    "y": self._read_int(str(y), "步骤 Y"),
                    "delay_ms": self._read_int(str(delay_ms), "步骤等待", 0),
                }
            )

        if not steps:
            raise ValueError("至少需要一个点击步骤")

        return {
            "window_title": self.window_title_var.get().strip(),
            "window_width": self._read_int(self.window_width_var.get(), "窗口宽度", 1),
            "window_height": self._read_int(self.window_height_var.get(), "窗口高度", 1),
            "window_x": self._read_int(self.window_x_var.get(), "窗口 X"),
            "window_y": self._read_int(self.window_y_var.get(), "窗口 Y"),
            "loop_forever": self.loop_forever_var.get(),
            "loop_count": self._read_int(self.loop_count_var.get(), "循环次数", 1),
            "startup_delay_ms": self._read_int(self.startup_delay_var.get(), "启动前等待", 0),
            "steps": steps,
        }

    def save_current_config(self) -> None:
        try:
            self.config_data = self.collect_config()
            save_config(self.config_data)
        except ValueError as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self.log("配置已保存到 config.json")

    def add_step(self) -> None:
        try:
            name = self.step_name_var.get().strip() or "未命名步骤"
            x = self._read_int(self.step_x_var.get(), "X")
            y = self._read_int(self.step_y_var.get(), "Y")
            delay_ms = self._read_int(self.step_delay_var.get(), "等待(ms)", 0)
        except ValueError as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        self.step_table.insert("", "end", values=(name, x, y, delay_ms))
        self.log(f"已新增步骤：{name} ({x}, {y})")

    def update_step(self) -> None:
        selection = self.step_table.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选中一个步骤")
            return

        try:
            name = self.step_name_var.get().strip() or "未命名步骤"
            x = self._read_int(self.step_x_var.get(), "X")
            y = self._read_int(self.step_y_var.get(), "Y")
            delay_ms = self._read_int(self.step_delay_var.get(), "等待(ms)", 0)
        except ValueError as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        self.step_table.item(selection[0], values=(name, x, y, delay_ms))
        self.log(f"已更新步骤：{name}")

    def delete_step(self) -> None:
        selection = self.step_table.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选中一个步骤")
            return

        for item in selection:
            self.step_table.delete(item)
        self.log("已删除选中的步骤")

    def move_step(self, direction: int) -> None:
        selection = self.step_table.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选中一个步骤")
            return

        item = selection[0]
        index = self.step_table.index(item)
        target = index + direction
        if target < 0 or target >= len(self.step_table.get_children()):
            return
        self.step_table.move(item, "", target)
        self.log("已调整步骤顺序")

    def run_once(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("提示", "当前已有任务在运行")
            return
        self._start_worker(single_run=True)

    def start_loop(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("提示", "循环已在运行中")
            return
        self._start_worker(single_run=False)

    def _start_worker(self, single_run: bool) -> None:
        try:
            self.config_data = self.collect_config()
            save_config(self.config_data)
        except ValueError as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self.stop_event.clear()
        self.worker_thread = threading.Thread(
            target=self._execute_flow,
            args=(self.config_data, single_run),
            daemon=True,
        )
        self.worker_thread.start()

    def stop_loop(self) -> None:
        self.stop_event.set()
        self.log("已请求停止，当前步骤结束后会退出")

    def _execute_flow(self, config: dict, single_run: bool) -> None:
        try:
            total_loops = 1 if single_run else config["loop_count"]
            current_loop = 0

            if config["startup_delay_ms"] > 0:
                self.log(f"启动等待 {config['startup_delay_ms']} ms")
                self._sleep_with_stop(config["startup_delay_ms"] / 1000)

            while not self.stop_event.is_set():
                current_loop += 1
                self.log(f"开始第 {current_loop} 轮")
                self._prepare_window(config)
                self._run_steps(config["steps"])

                if single_run:
                    break
                if not config["loop_forever"] and current_loop >= total_loops:
                    break

            if self.stop_event.is_set():
                self.log("任务已停止")
            else:
                self.log("任务执行完成")
        except Exception as exc:  # noqa: BLE001
            self.log(f"运行失败：{exc}")
            self.root.after(0, lambda: messagebox.showerror("运行失败", str(exc)))

    def _prepare_window(self, config: dict) -> None:
        title = config["window_title"]
        if not title:
            raise RuntimeError("窗口标题不能为空")

        hwnd = find_window(title)
        if not hwnd:
            raise RuntimeError(f"未找到窗口：{title}")

        activate_window(hwnd)
        resize_and_move_window(
            hwnd,
            config["window_x"],
            config["window_y"],
            config["window_width"],
            config["window_height"],
        )
        self.log(f"已激活窗口并调整大小：{title}")
        time.sleep(0.5)

    def _run_steps(self, steps: list[dict]) -> None:
        screen_width = user32.GetSystemMetrics(SM_CXSCREEN)
        screen_height = user32.GetSystemMetrics(SM_CYSCREEN)

        for step in steps:
            if self.stop_event.is_set():
                return

            x = int(step["x"])
            y = int(step["y"])
            if x < 0 or y < 0 or x >= screen_width or y >= screen_height:
                raise RuntimeError(f"坐标超出屏幕范围：{step['name']} ({x}, {y})")

            self.log(f"点击：{step['name']} ({x}, {y})")
            click_screen(x, y)
            self._sleep_with_stop(int(step["delay_ms"]) / 1000)

    def _sleep_with_stop(self, seconds: float) -> None:
        end_time = time.time() + seconds
        while time.time() < end_time:
            if self.stop_event.is_set():
                return
            time.sleep(0.1)

    def on_close(self) -> None:
        self.stop_event.set()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = AutoClickApp(root)
    app.log("应用已启动，默认配置已载入")
    root.mainloop()


if __name__ == "__main__":
    main()
