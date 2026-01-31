import sys
import os
import time
import json
import pyautogui
import pyperclip
import traceback
from typing import Callable, Optional
import subprocess
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QPushButton, QLabel, QComboBox, QLineEdit, QScrollArea, 
                               QFileDialog, QTextEdit, QMessageBox, QFrame)
from PySide6.QtCore import Qt, QThread, Signal

# --------------------------
# 核心逻辑 (原 waterRPA.py)
# --------------------------

def _get_frontmost_app_name() -> Optional[str]:
    """macOS: 获取最前台应用名（用于判断是否只是激活窗口而未触发控件）。"""
    if not _is_macos():
        return None
    try:
        res = subprocess.run(
            [
                "osascript",
                "-e",
                "tell application \"System Events\" to get name of first application process whose frontmost is true",
            ],
            capture_output=True,
            text=True,
            timeout=1,
        )
        if res.returncode != 0:
            return None
        name = (res.stdout or "").strip()
        return name or None
    except Exception:
        return None


class TaskStopped(Exception):
    """用户请求停止任务（可取消执行）。"""


class StepFailed(Exception):
    """某个步骤执行失败（按策略：失败即停止整个任务）。"""


def _is_macos() -> bool:
    return sys.platform == "darwin"


def _cancellable_sleep(
    seconds: float,
    should_stop: Optional[Callable[[], bool]] = None,
    tick: float = 0.1,
):
    """分片 sleep，确保 stop 能在 tick 粒度内生效。"""
    if seconds <= 0:
        return
    if tick <= 0:
        tick = 0.1

    end_time = time.time() + seconds
    while True:
        if should_stop and should_stop():
            raise TaskStopped("任务已停止")

        remaining = end_time - time.time()
        if remaining <= 0:
            return

        time.sleep(min(tick, remaining))


def _normalize_xy_for_macos_retina(
    x: float,
    y: float,
    *,
    scale_x: Optional[float],
    scale_y: Optional[float],
) -> tuple:
    """
    macOS Retina 常见现象：截图像素坐标是屏幕点坐标的 2 倍。
    若检测到 scale != 1，则把 locate 的坐标缩放回屏幕坐标。
    """
    if not _is_macos():
        return x, y
    if not scale_x or not scale_y:
        return x, y
    if abs(scale_x - 1.0) < 0.01 and abs(scale_y - 1.0) < 0.01:
        return x, y

    return x / scale_x, y / scale_y


def _locate_center_on_screen(
    img: str,
    *,
    confidence: float = 0.9,
    on_warn: Optional[Callable[[str], None]] = None,
):
    """
    兼容 OpenCV/置信度参数不可用的场景：
    - 优先使用 confidence（更稳定）
    - 若环境不支持（常见：未安装 OpenCV），降级为不带 confidence 的匹配并提示一次
    """
    if not img:
        return None

    # 如果给的是文件路径（绝对或相对），尽早发现明显的文件缺失问题
    # 允许用户传入非文件路径（例如某些自定义方式），此时不做存在性校验
    if any(sep in img for sep in ("/", "\\")) and not os.path.exists(img):
        raise StepFailed(f"图片文件不存在: {img}")

    try:
        return pyautogui.locateCenterOnScreen(img, confidence=confidence)
    except Exception as e:
        # 典型报错：confidence 参数仅在安装 OpenCV 时可用
        msg = str(e)
        if "confidence" in msg.lower() or "opencv" in msg.lower():
            if on_warn:
                on_warn("检测到环境不支持 confidence/OpenCV，已降级为不带置信度的找图。建议安装 opencv-python 提升稳定性。")
            try:
                return pyautogui.locateCenterOnScreen(img)
            except pyautogui.ImageNotFoundException:
                return None
        raise


def mouseClick(
    clickTimes,
    lOrR,
    img,
    reTry,
    timeout=60,
    should_stop: Optional[Callable[[], bool]] = None,
    on_warn: Optional[Callable[[str], None]] = None,
    scale_x: Optional[float] = None,
    scale_y: Optional[float] = None,
):
    """
    安全统一语义：reTry 仅代表“找图重试策略”，不承载“重复点击”语义。

    reTry:
      - 1: 只尝试一次，找不到则失败
      - >1: 最多尝试 N 次，找到则点击一次并继续
      - -1: 无限等待直到首次匹配成功，点击一次并继续

    timeout: 超时时间(秒)，默认60秒。防止无限卡死。
    """
    start_time = time.time()

    # 规范化 reTry
    try:
        reTry = int(reTry)
    except Exception:
        reTry = 1

    if reTry == 0 or reTry < -1:
        reTry = 1

    def _check_timeout():
        if timeout and (time.time() - start_time > timeout):
            raise StepFailed(f"等待图片超时 ({timeout}秒): {img}")

    # reTry=-1：无限等待直到成功（但仍受 timeout/stop 影响）
    if reTry == -1:
        while True:
            if should_stop and should_stop():
                raise TaskStopped("任务已停止")
            _check_timeout()

            try:
                location = _locate_center_on_screen(img, on_warn=on_warn)
            except pyautogui.ImageNotFoundException:
                location = None

            if location is not None:
                nx, ny = _normalize_xy_for_macos_retina(
                    location.x,
                    location.y,
                    scale_x=scale_x,
                    scale_y=scale_y,
                )
                pre_app = _get_frontmost_app_name()
                pyautogui.click(
                    int(round(nx)),
                    int(round(ny)),
                    clicks=clickTimes,
                    interval=0.2,
                    duration=0.2,
                    button=lOrR,
                )
                post_app = _get_frontmost_app_name()
                # 如果本次点击导致前台应用切换，常见现象是“第一次点击只激活窗口”
                # 这里做一次无延迟的补偿点击，尽量让控件动作生效。
                if pre_app and post_app and pre_app != post_app:
                    pyautogui.click(
                        int(round(nx)),
                        int(round(ny)),
                        clicks=clickTimes,
                        interval=0.2,
                        duration=0.2,
                        button=lOrR,
                    )
                return

            _cancellable_sleep(0.1, should_stop)

    # reTry>=1：有限次尝试
    for attempt in range(reTry):
        if should_stop and should_stop():
            raise TaskStopped("任务已停止")
        _check_timeout()

        try:
            location = _locate_center_on_screen(img, on_warn=on_warn)
        except pyautogui.ImageNotFoundException:
            location = None

        if location is not None:
            nx, ny = _normalize_xy_for_macos_retina(
                location.x,
                location.y,
                scale_x=scale_x,
                scale_y=scale_y,
            )
            pre_app = _get_frontmost_app_name()
            pyautogui.click(
                int(round(nx)),
                int(round(ny)),
                clicks=clickTimes,
                interval=0.2,
                duration=0.2,
                button=lOrR,
            )
            post_app = _get_frontmost_app_name()
            if pre_app and post_app and pre_app != post_app:
                pyautogui.click(
                    int(round(nx)),
                    int(round(ny)),
                    clicks=clickTimes,
                    interval=0.2,
                    duration=0.2,
                    button=lOrR,
                )
            return

        if attempt < reTry - 1:
            _cancellable_sleep(0.1, should_stop)

    raise StepFailed(f"未找到匹配图片: {img}")


def mouseMove(
    img,
    reTry,
    timeout=60,
    should_stop: Optional[Callable[[], bool]] = None,
    on_warn: Optional[Callable[[str], None]] = None,
    scale_x: Optional[float] = None,
    scale_y: Optional[float] = None,
):
    """
    鼠标悬停（移动但不点击）
    """
    start_time = time.time()
    try:
        reTry = int(reTry)
    except Exception:
        reTry = 1

    if reTry == 0 or reTry < -1:
        reTry = 1

    def _check_timeout():
        if timeout and (time.time() - start_time > timeout):
            raise StepFailed(f"等待图片超时 ({timeout}秒): {img}")

    if reTry == -1:
        while True:
            if should_stop and should_stop():
                raise TaskStopped("任务已停止")
            _check_timeout()

            try:
                location = _locate_center_on_screen(img, on_warn=on_warn)
            except pyautogui.ImageNotFoundException:
                location = None

            if location is not None:
                nx, ny = _normalize_xy_for_macos_retina(
                    location.x,
                    location.y,
                    scale_x=scale_x,
                    scale_y=scale_y,
                )
                pyautogui.moveTo(int(round(nx)), int(round(ny)), duration=0.2)
                return

            _cancellable_sleep(0.1, should_stop)

    for attempt in range(reTry):
        if should_stop and should_stop():
            raise TaskStopped("任务已停止")
        _check_timeout()

        try:
            location = _locate_center_on_screen(img, on_warn=on_warn)
        except pyautogui.ImageNotFoundException:
            location = None

        if location is not None:
            nx, ny = _normalize_xy_for_macos_retina(
                location.x,
                location.y,
                scale_x=scale_x,
                scale_y=scale_y,
            )
            pyautogui.moveTo(int(round(nx)), int(round(ny)), duration=0.2)
            return

        if attempt < reTry - 1:
            _cancellable_sleep(0.1, should_stop)

    raise StepFailed(f"未找到匹配图片: {img}")

class RPAEngine:
    def __init__(self):
        self.is_running = False
        self.stop_requested = False

    def stop(self):
        self.stop_requested = True
        self.is_running = False

    def run_tasks(self, tasks, loop_forever=False, callback_msg=None):
        """
        tasks: list of dict, format:
        [
            {"type": 1.0, "value": "1.png", "retry": 1},
            ...
        ]
        """
        self.is_running = True
        self.stop_requested = False

        def should_stop() -> bool:
            return bool(self.stop_requested)

        # 用于输出一次性的降级提示（例如 OpenCV 不可用）
        warned_messages = set()

        def warn_once(msg: str):
            if msg in warned_messages:
                return
            warned_messages.add(msg)
            if callback_msg:
                callback_msg(f"提示: {msg}")

        try:
            screen_w, screen_h = pyautogui.size()
        except Exception:
            screen_w, screen_h = None, None
        try:
            shot = pyautogui.screenshot()
            shot_w, shot_h = shot.size
        except Exception:
            shot_w, shot_h = None, None

        scale_x = (shot_w / screen_w) if (shot_w and screen_w) else None
        scale_y = (shot_h / screen_h) if (shot_h and screen_h) else None

        try:
            while True:
                for idx, task in enumerate(tasks):
                    if self.stop_requested:
                        if callback_msg: callback_msg("任务已停止")
                        return

                    cmd_type = task.get("type")
                    cmd_value = task.get("value")
                    retry = task.get("retry", 1)

                    if callback_msg:
                        callback_msg(f"执行步骤 {idx+1}: 类型={cmd_type}, 内容={cmd_value}")

                    try:
                        if cmd_type == 1.0: # 单击左键
                            mouseClick(
                                1,
                                "left",
                                cmd_value,
                                retry,
                                should_stop=should_stop,
                                on_warn=warn_once,
                                scale_x=scale_x,
                                scale_y=scale_y,
                            )
                            if callback_msg: callback_msg(f"单击左键: {cmd_value}")
                        
                        elif cmd_type == 2.0: # 双击左键
                            mouseClick(
                                2,
                                "left",
                                cmd_value,
                                retry,
                                should_stop=should_stop,
                                on_warn=warn_once,
                                scale_x=scale_x,
                                scale_y=scale_y,
                            )
                            if callback_msg: callback_msg(f"双击左键: {cmd_value}")
                        
                        elif cmd_type == 3.0: # 右键
                            mouseClick(
                                1,
                                "right",
                                cmd_value,
                                retry,
                                should_stop=should_stop,
                                on_warn=warn_once,
                                scale_x=scale_x,
                                scale_y=scale_y,
                            )
                            if callback_msg: callback_msg(f"右键单击: {cmd_value}")
                        
                        elif cmd_type == 4.0: # 输入
                            pyperclip.copy(str(cmd_value))
                            if _is_macos():
                                pyautogui.hotkey("command", "v")
                            else:
                                pyautogui.hotkey("ctrl", "v")
                            _cancellable_sleep(0.5, should_stop)
                            if callback_msg: callback_msg(f"输入文本: {cmd_value}")
                        
                        elif cmd_type == 5.0: # 等待
                            sleep_time = float(cmd_value)
                            _cancellable_sleep(sleep_time, should_stop)
                            if callback_msg: callback_msg(f"等待 {sleep_time} 秒")
                        
                        elif cmd_type == 6.0: # 滚轮
                            scroll_val = int(cmd_value)
                            pyautogui.scroll(scroll_val)
                            if callback_msg: callback_msg(f"滚轮滑动 {scroll_val}")

                        elif cmd_type == 7.0: # 系统按键 (组合键)
                            keys = str(cmd_value).lower().split('+')
                            # 去除空格
                            keys = [k.strip() for k in keys]
                            # 轻量兼容别名：cmd/control/option 等
                            key_alias = {
                                "cmd": "command",
                                "command": "command",
                                "ctl": "ctrl",
                                "control": "ctrl",
                                "option": "alt",
                                "win": "winleft",
                                "windows": "winleft",
                                "super": "winleft",
                            }
                            keys = [key_alias.get(k, k) for k in keys if k]
                            pyautogui.hotkey(*keys)
                            if callback_msg: callback_msg(f"按键组合: {cmd_value}")

                        elif cmd_type == 8.0: # 鼠标悬停
                            mouseMove(
                                cmd_value,
                                retry,
                                should_stop=should_stop,
                                on_warn=warn_once,
                                scale_x=scale_x,
                                scale_y=scale_y,
                            )
                            if callback_msg: callback_msg(f"鼠标悬停: {cmd_value}")

                        elif cmd_type == 9.0: # 截图保存
                            path = str(cmd_value)
                            # 如果是目录，自动拼接时间戳文件名
                            if os.path.isdir(path):
                                timestamp = time.strftime("%Y%m%d_%H%M%S")
                                filename = os.path.join(path, f"screenshot_{timestamp}.png")
                            else:
                                # 兼容旧逻辑：如果用户直接输入了带文件名的路径
                                filename = path
                                if not filename.endswith(('.png', '.jpg', '.bmp')):
                                    filename += '.png'
                            
                            pyautogui.screenshot(filename)
                            if callback_msg: callback_msg(f"截图已保存: {filename}")
                        else:
                            raise StepFailed(f"未知指令类型: {cmd_type}")

                    except StepFailed as e:
                        if callback_msg:
                            callback_msg(f"步骤 {idx+1} 失败: 类型={cmd_type}, 内容={cmd_value}, 原因={e}")
                        return

                if not loop_forever:
                    break
                
                if callback_msg: callback_msg("等待 0.1 秒进入下一轮循环...")
                _cancellable_sleep(0.1, should_stop)
                
        except TaskStopped:
            if callback_msg:
                callback_msg("任务已停止")
        except Exception as e:
            if callback_msg: callback_msg(f"执行出错: {e}")
            traceback.print_exc()
        finally:
            self.is_running = False
            if callback_msg: callback_msg("任务结束")

# --------------------------
# GUI 界面 (原 rpa_gui.py)
# --------------------------

# 定义操作类型映射
CMD_TYPES = {
    "左键单击": 1.0,
    "左键双击": 2.0,
    "右键单击": 3.0,
    "输入文本": 4.0,
    "等待(秒)": 5.0,
    "滚轮滑动": 6.0,
    "系统按键": 7.0,
    "鼠标悬停": 8.0,
    "截图保存": 9.0
}

CMD_TYPES_REV = {v: k for k, v in CMD_TYPES.items()}

class WorkerThread(QThread):
    log_signal = Signal(str)
    finished_signal = Signal()

    def __init__(self, engine, tasks, loop_forever):
        super().__init__()
        self.engine = engine
        self.tasks = tasks
        self.loop_forever = loop_forever

    def run(self):
        self.engine.run_tasks(self.tasks, self.loop_forever, self.log_callback)
        self.finished_signal.emit()

    def log_callback(self, msg):
        self.log_signal.emit(msg)

class TaskRow(QFrame):
    def __init__(self, parent_layout, delete_callback):
        super().__init__()
        self.setFrameShape(QFrame.StyledPanel)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(5, 5, 5, 5)
        
        # 操作类型选择
        self.type_combo = QComboBox()
        self.type_combo.addItems(list(CMD_TYPES.keys()))
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        self.layout.addWidget(self.type_combo)
        
        # 参数输入区域
        self.value_input = QLineEdit()
        self.value_input.setPlaceholderText("参数值 (如图片路径、文本、时间)")
        self.layout.addWidget(self.value_input)
        
        # 文件选择按钮 (默认隐藏)
        self.file_btn = QPushButton("选择图片")
        self.file_btn.clicked.connect(self.select_file)
        self.file_btn.setVisible(True) # 默认是左键单击，需要显示
        self.layout.addWidget(self.file_btn)
        
        # 重试次数 (默认隐藏)
        self.retry_input = QLineEdit()
        self.retry_input.setPlaceholderText("重试次数 (1=一次, -1=无限)")
        self.retry_input.setText("1")
        self.retry_input.setFixedWidth(100)
        self.retry_input.setVisible(True)
        self.layout.addWidget(self.retry_input)
        
        # 删除按钮
        self.del_btn = QPushButton("X")
        self.del_btn.setStyleSheet("color: red; font-weight: bold;")
        self.del_btn.setFixedWidth(30)
        self.del_btn.clicked.connect(lambda: delete_callback(self))
        self.layout.addWidget(self.del_btn)
        
        parent_layout.addWidget(self)

    def on_type_changed(self, text):
        cmd_type = CMD_TYPES[text]
        
        # 图片相关操作 (1, 2, 3, 8)
        if cmd_type in [1.0, 2.0, 3.0, 8.0]:
            self.file_btn.setVisible(True)
            self.file_btn.setText("选择图片")
            self.retry_input.setVisible(True)
            self.value_input.setPlaceholderText("图片路径")
        # 输入 (4)
        elif cmd_type == 4.0:
            self.file_btn.setVisible(False)
            self.retry_input.setVisible(False)
            self.value_input.setPlaceholderText("请输入要发送的文本")
        # 等待 (5)
        elif cmd_type == 5.0:
            self.file_btn.setVisible(False)
            self.retry_input.setVisible(False)
            self.value_input.setPlaceholderText("等待秒数 (如 1.5)")
        # 滚轮 (6)
        elif cmd_type == 6.0:
            self.file_btn.setVisible(False)
            self.retry_input.setVisible(False)
            self.value_input.setPlaceholderText("滚动距离 (正数向上，负数向下)")
        # 系统按键 (7)
        elif cmd_type == 7.0:
            self.file_btn.setVisible(False)
            self.retry_input.setVisible(False)
            self.value_input.setPlaceholderText("组合键 (如 ctrl+s, alt+tab)")
        # 截图保存 (9)
        elif cmd_type == 9.0:
            self.file_btn.setVisible(True)
            self.file_btn.setText("选择保存文件夹")
            self.retry_input.setVisible(False)
            self.value_input.setPlaceholderText("保存目录 (如 D:\\Screenshots)")

    def set_data(self, data):
        """用于回填数据"""
        cmd_type = data.get("type")
        value = data.get("value", "")
        retry = data.get("retry", 1)

        # 设置类型 (反向查找文本)
        if cmd_type in CMD_TYPES_REV:
            self.type_combo.setCurrentText(CMD_TYPES_REV[cmd_type])
        
        # 设置值
        self.value_input.setText(str(value))
        
        # 设置重试次数
        self.retry_input.setText(str(retry))

    def select_file(self):
        cmd_type = CMD_TYPES[self.type_combo.currentText()]
        
        # 截图保存 (9.0) -> 选择文件夹
        if cmd_type == 9.0:
            folder = QFileDialog.getExistingDirectory(self, "选择保存文件夹", os.getcwd())
            if folder:
                self.value_input.setText(folder)
        
        # 其他图片操作 (1, 2, 3, 8) -> 打开文件对话框
        else:
            filename, _ = QFileDialog.getOpenFileName(self, "选择图片", os.getcwd(), "Image Files (*.png *.jpg *.bmp)")
            if filename:
                self.value_input.setText(filename)

    def get_data(self):
        cmd_type = CMD_TYPES[self.type_combo.currentText()]
        value = self.value_input.text()
        
        # 数据校验与转换
        try:
            if cmd_type in [5.0, 6.0]:
                # 尝试转换为数字，如果失败可能会在运行时报错，这里简单处理
                if not value: value = "0"
            
            retry = 1
            if self.retry_input.isVisible():
                retry_text = self.retry_input.text()
                if retry_text:
                    retry = int(retry_text)
        except ValueError:
            pass # 保持默认

        return {
            "type": cmd_type,
            "value": value,
            "retry": retry
        }

class RPAWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("不高兴就喝水 RPA 配置工具")
        self.resize(800, 600)
        
        self.engine = RPAEngine()
        self.worker = None
        self.rows = []

        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 顶部控制栏
        top_bar = QHBoxLayout()
        
        self.add_btn = QPushButton("+ 新增指令")
        self.add_btn.clicked.connect(self.add_row)
        top_bar.addWidget(self.add_btn)

        self.save_btn = QPushButton("保存配置")
        self.save_btn.clicked.connect(self.save_config)
        top_bar.addWidget(self.save_btn)

        self.load_btn = QPushButton("导入配置")
        self.load_btn.clicked.connect(self.load_config)
        top_bar.addWidget(self.load_btn)
        
        top_bar.addStretch()
        
        self.loop_check = QComboBox()
        self.loop_check.addItems(["执行一次", "循环执行"])
        top_bar.addWidget(self.loop_check)
        
        self.start_btn = QPushButton("开始运行")
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        self.start_btn.clicked.connect(self.start_task)
        top_bar.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setStyleSheet("background-color: #f44336; color: white;")
        self.stop_btn.clicked.connect(self.stop_task)
        self.stop_btn.setEnabled(False)
        top_bar.addWidget(self.stop_btn)
        
        main_layout.addLayout(top_bar)

        # 任务列表区域 (滚动)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.task_container = QWidget()
        self.task_layout = QVBoxLayout(self.task_container)
        self.task_layout.addStretch() # 弹簧，确保添加的行在顶部
        scroll.setWidget(self.task_container)
        main_layout.addWidget(scroll)

        # 日志区域
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(150)
        main_layout.addWidget(QLabel("运行日志:"))
        main_layout.addWidget(self.log_area)

        # 初始添加一行
        self.add_row()

    def add_row(self, data=None):
        # 移除底部的弹簧
        self.task_layout.takeAt(self.task_layout.count() - 1)
        
        row = TaskRow(self.task_layout, self.delete_row)
        if data:
            row.set_data(data)
        self.rows.append(row)
        
        # 加回弹簧
        self.task_layout.addStretch()

    def delete_row(self, row_widget):
        if row_widget in self.rows:
            self.rows.remove(row_widget)
            row_widget.deleteLater()
            
    def save_config(self):
        tasks = []
        for row in self.rows:
            data = row.get_data()
            # 允许保存空值，方便后续编辑
            tasks.append(data)
            
        if not tasks:
            QMessageBox.warning(self, "警告", "没有可保存的配置")
            return

        filename, _ = QFileDialog.getSaveFileName(self, "保存配置", os.getcwd(), "JSON Files (*.json);;Text Files (*.txt)")
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(tasks, f, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "成功", "配置已保存！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存失败: {e}")

    def load_config(self):
        filename, _ = QFileDialog.getOpenFileName(self, "导入配置", os.getcwd(), "JSON Files (*.json);;Text Files (*.txt)")
        if not filename:
            return
            
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                tasks = json.load(f)
            
            if not isinstance(tasks, list):
                raise ValueError("文件格式不正确")

            # 清空现有行
            for row in self.rows:
                row.deleteLater()
            self.rows.clear()
            
            # 重新添加行
            for task in tasks:
                self.add_row(task)
                
            QMessageBox.information(self, "成功", f"成功导入 {len(tasks)} 条指令！")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入失败: {e}")

    def start_task(self):
        tasks = []
        for row in self.rows:
            data = row.get_data()
            if not data['value']:
                QMessageBox.warning(self, "警告", "请检查有空参数的指令！")
                return
            tasks.append(data)
            
        if not tasks:
            QMessageBox.warning(self, "警告", "请至少添加一条指令！")
            return

        self.log_area.clear()
        self.log("任务开始...")
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.add_btn.setEnabled(False)
        
        loop = (self.loop_check.currentText() == "循环执行")
        
        self.worker = WorkerThread(self.engine, tasks, loop)
        self.worker.log_signal.connect(self.log)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

        # 最小化窗口
        self.showMinimized()

    def stop_task(self):
        self.engine.stop()
        self.log("正在停止...")

    def on_finished(self):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.add_btn.setEnabled(True)
        self.log("任务已结束")
        
        # 恢复窗口并置顶
        self.showNormal()
        self.activateWindow()

    def log(self, msg):
        self.log_area.append(msg)

    def closeEvent(self, event):
        """窗口关闭事件：确保线程停止，防止残留"""
        if self.worker and self.worker.isRunning():
            self.engine.stop()
            self.worker.quit()
            self.worker.wait()
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = RPAWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
