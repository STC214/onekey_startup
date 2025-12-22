import sys
import os
import json
import subprocess
import ctypes
import time
import re
import base64
from pathlib import Path

# 检查是否安装了必要的包，如果没有则尝试导入备用模块
try:
    import psutil
except ImportError:
    psutil = None

try:
    from win32com.client import Dispatch
    import pywintypes
except ImportError:
    Dispatch = None
    pywintypes = None

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, 
    QFileDialog, QSystemTrayIcon, QMenu, QAction, QScrollArea, QFrame, QSizePolicy, QMessageBox, 
    QAbstractItemView, QTableWidget, QTableWidgetItem, QHeaderView, QStyle, QStyleOptionButton, 
    QCheckBox, QDialog
)
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal, QTimer, QPoint, QRect
from PyQt5.QtGui import QIcon, QPalette, QColor, QFont, QPainter, QBrush, QPen, QPixmap

# 内嵌的图标数据 (base64编码的ICO文件)
APP_ICON_DATA = """AAABAAEAEBAAAAAAAABoBQAAFgAAACgAAAAQAAAAIAAAAAEACAAAAAAAAAEAAAAAAAAAAAAAAAEAAAAAAAABAAAAACAAAAAEAAEAAAAAAAEAEAAAAAAQAAAQAAAAAAAAEAAAAAAAAAAAAAAAAP//AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A"""


# 检查管理员权限
def is_admin():
    """检查当前是否具有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员权限重新运行当前程序"""
    script = os.path.abspath(sys.argv[0])
    params = ' '.join([script] + sys.argv[1:])
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        return True
    except:
        return False

def get_app_icon():
    """获取应用图标，支持打包后的exe和开发环境"""
    try:
        # 尝试从资源中加载图标
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller打包后的路径
            icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            if os.path.exists(icon_path):
                return QIcon(icon_path)
        
        # 尝试加载当前目录下的图标
        if os.path.exists("icon.ico"):
            return QIcon("icon.ico")
        
        # 使用内嵌的图标数据
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(APP_ICON_DATA))
        if not pixmap.isNull():
            return QIcon(pixmap)
    except Exception as e:
        print(f"加载图标时出错: {e}")
    
    # 如果所有方法都失败，返回默认图标
    return QApplication.style().standardIcon(QStyle.SP_ComputerIcon)

# Darcula主题调色板
class DarculaPalette:
    BACKGROUND = QColor(43, 43, 43)
    FOREGROUND = QColor(169, 183, 198)
    SELECTION = QColor(60, 100, 150)
    BORDER = QColor(75, 75, 75)
    BUTTON = QColor(60, 60, 60)
    BUTTON_HOVER = QColor(70, 70, 70)
    BUTTON_PRESSED = QColor(50, 50, 50)
    INPUT_BG = QColor(55, 55, 55)
    INPUT_TEXT = QColor(192, 192, 192)
    DISABLED = QColor(90, 90, 90)
    TITLE_BAR = QColor(30, 30, 30)
    SUCCESS = QColor(80, 160, 80)
    ERROR = QColor(190, 80, 80)

# 标题栏按钮
class TitleBarButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFixedSize(45, 30)
        self.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                color: #b0b0b0;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #404040;
            }
            QPushButton:pressed {
                background-color: #303030;
            }
        """)

# 程序管理行
class ProgramRow(QWidget):
    def __init__(self, manager=None, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.process_name = None
        self.is_uwp = False
        self.selected_process = None
        self.running = False
        
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        # 程序路径输入框
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("拖放程序或快捷方式，或点击浏览按钮")
        self.path_input.setStyleSheet("""
            QLineEdit {
                background-color: #373737;
                border: 1px solid #4B4B4B;
                color: #c0c0c0;
                border-radius: 3px;
                padding: 5px;
                min-height: 25px;
                height: 25px;
                font-size: 10pt;
            }
            QLineEdit:focus {
                border: 1px solid #3C6496;
            }
        """)
        self.path_input.setAcceptDrops(True)
        layout.addWidget(self.path_input, 4)
        
        # 浏览按钮
        self.browse_btn = QPushButton("浏览")
        self.browse_btn.setMinimumSize(60, 25)
        self.browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #3C3C3C;
                color: #a9b7c6;
                border: 1px solid #4B4B4B;
                border-radius: 3px;
                padding: 3px;
            }
            QPushButton:hover {
                background-color: #484848;
            }
            QPushButton:pressed {
                background-color: #323232;
            }
        """)
        self.browse_btn.clicked.connect(self.browse_file)
        layout.addWidget(self.browse_btn, 1)
        
        # 状态显示 - 与输入框高度保持一致
        self.status_label = QLabel("未运行")
        self.status_label.setFixedHeight(25)  # 设置固定高度
        self.status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # 禁止垂直拉伸
        self.status_label.setAlignment(Qt.AlignCenter)  # 文本居中
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #373737;
                border: 1px solid #4B4B4B;
                color: #808080;
                border-radius: 3px;
                padding: 3px 5px;
                font-size: 10pt;
            }
        """)
        layout.addWidget(self.status_label, 2)
        
        # 选择进程按钮
        self.select_process_btn = QPushButton("选择进程")
        self.select_process_btn.setMinimumSize(80, 25)
        self.select_process_btn.setStyleSheet("""
            QPushButton {
                background-color: #3C3C3C;
                color: #a9b7c6;
                border: 1px solid #4B4B4B;
                border-radius: 3px;
                padding: 3px;
            }
            QPushButton:hover {
                background-color: #484848;
            }
            QPushButton:pressed {
                background-color: #323232;
            }
        """)
        self.select_process_btn.clicked.connect(self.select_process)
        layout.addWidget(self.select_process_btn, 1)
        
        # 删除按钮
        self.delete_btn = QPushButton("×")
        self.delete_btn.setMinimumSize(25, 25)
        self.delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #3C3C3C;
                color: #a9b7c6;
                border: 1px solid #4B4B4B;
                border-radius: 3px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #503030;
                color: white;
            }
            QPushButton:pressed {
                background-color: #323232;
            }
        """)
        self.delete_btn.clicked.connect(self.delete_row)
        layout.addWidget(self.delete_btn, 1)
        
        self.setLayout(layout)
        self.setAcceptDrops(True)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.path_input.setText(files[0])
            self.check_if_uwp(files[0])
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择程序或快捷方式",
            "",
            "可执行文件和快捷方式 (*.exe *.lnk *.bat *.cmd);;所有文件 (*.*)"
        )
        if file_path:
            self.path_input.setText(file_path)
            self.check_if_uwp(file_path)
    
    def check_if_uwp(self, file_path):
        """检查是否为UWP应用快捷方式"""
        self.is_uwp = False
        if file_path.lower().endswith('.lnk') and Dispatch:
            try:
                shell = Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(file_path)
                target_path = shortcut.TargetPath
                # UWP应用通常有AppX标记或没有.exe扩展名
                if "AppX" in target_path or not target_path.lower().endswith('.exe'):
                    self.is_uwp = True
                    # 从快捷方式文件名提取UWP应用名称
                    app_name = os.path.splitext(os.path.basename(file_path))[0]
                    self.process_name = app_name
            except:
                pass
    
    def set_status(self, running, process_name=None):
        self.running = running
        if running:
            self.status_label.setText(process_name or "运行中")
            self.status_label.setStyleSheet("""
                QLabel {
                    background-color: #373737;
                    border: 1px solid #4B4B4B;
                    color: #50A050;
                    border-radius: 3px;
                    padding: 3px 5px;
                    font-size: 10pt;
                }
            """)
        else:
            self.status_label.setText("未运行")
            self.status_label.setStyleSheet("""
                QLabel {
                    background-color: #373737;
                    border: 1px solid #4B4B4B;
                    color: #808080;
                    border-radius: 3px;
                    padding: 3px 5px;
                    font-size: 10pt;
                }
            """)
    
    def select_process(self):
        if not self.manager:
            return
        dialog = ProcessSelectorDialog(self)
        if dialog.exec_():
            self.selected_process = dialog.selected_process
            if self.selected_process:
                self.process_name = self.selected_process
                self.status_label.setText(f"已选择: {self.selected_process}")
    
    def delete_row(self):
        if self.manager:
            self.manager.remove_program_row(self)
    
    def get_program_path(self):
        return self.path_input.text().strip()
    
    def is_valid(self):
        path = self.get_program_path()
        return bool(path) and os.path.exists(path)

# 进程选择对话框
class ProcessSelectorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择进程")
        self.setGeometry(100, 100, 600, 400)
        self.selected_process = None
        
        layout = QVBoxLayout()
        
        # 搜索框
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("搜索进程...")
        self.search_box.textChanged.connect(self.filter_processes)
        layout.addWidget(self.search_box)
        
        # 进程列表
        self.process_table = QTableWidget()
        self.process_table.setColumnCount(3)
        self.process_table.setHorizontalHeaderLabels(["进程名", "PID", "路径"])
        self.process_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.process_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.process_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.process_table.doubleClicked.connect(self.accept_selection)
        
        # 调整列宽
        header = self.process_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        
        layout.addWidget(self.process_table)
        
        # 按钮
        button_layout = QHBoxLayout()
        self.select_btn = QPushButton("选择")
        self.select_btn.clicked.connect(self.accept_selection)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.select_btn)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
        # 加载进程
        self.load_processes()
        
        # 应用样式
        self.setStyleSheet("""
            QDialog {
                background-color: #2B2B2B;
                color: #A9B7C6;
            }
            QLineEdit {
                background-color: #373737;
                border: 1px solid #4B4B4B;
                color: #c0c0c0;
                border-radius: 3px;
                padding: 5px;
            }
            QTableWidget {
                background-color: #373737;
                border: 1px solid #4B4B4B;
                color: #c0c0c0;
                gridline-color: #4B4B4B;
            }
            QTableWidget::item:selected {
                background-color: #3C6496;
                color: white;
            }
            QHeaderView::section {
                background-color: #303030;
                color: #a9b7c6;
                padding: 4px;
                border: 1px solid #4B4B4B;
            }
            QPushButton {
                background-color: #3C3C3C;
                color: #a9b7c6;
                border: 1px solid #4B4B4B;
                border-radius: 3px;
                padding: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #484848;
            }
            QPushButton:pressed {
                background-color: #323232;
            }
        """)
    
    def load_processes(self):
        self.all_processes = []
        if psutil:
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    info = proc.info
                    if info['exe']:
                        self.all_processes.append((
                            info['name'],
                            str(info['pid']),
                            info['exe']
                        ))
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
        
        self.process_table.setRowCount(len(self.all_processes))
        for row, proc_info in enumerate(self.all_processes):
            for col, item in enumerate(proc_info):
                table_item = QTableWidgetItem(item)
                table_item.setFlags(table_item.flags() & ~Qt.ItemIsEditable)
                self.process_table.setItem(row, col, table_item)
    
    def filter_processes(self, text):
        text = text.lower()
        for row in range(self.process_table.rowCount()):
            name_item = self.process_table.item(row, 0)
            path_item = self.process_table.item(row, 2)
            show = False
            if name_item and text in name_item.text().lower():
                show = True
            if path_item and text in path_item.text().lower():
                show = True
            self.process_table.setRowHidden(row, not show)
    
    def accept_selection(self):
        selected_items = self.process_table.selectedItems()
        if selected_items:
            row = selected_items[0].row()
            self.selected_process = self.process_table.item(row, 0).text()
            self.accept()
        else:
            QMessageBox.warning(self, "警告", "请先选择一个进程")

# 自定义标题栏
class CustomTitleBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(QPalette.Window, DarculaPalette.TITLE_BAR)
        self.setPalette(palette)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(10, 0, 0, 0)
        layout.setSpacing(0)
        
        # 应用标题
        self.title = QLabel("程序启动管理器")
        self.title.setStyleSheet("color: #a9b7c6; font-weight: bold; font-size: 10pt;")
        layout.addWidget(self.title)
        layout.addStretch()
        
        # 最小化按钮
        self.minimize_btn = TitleBarButton("–")
        self.minimize_btn.clicked.connect(self.minimize_window)
        layout.addWidget(self.minimize_btn)
        
        # 最大化/还原按钮
        self.maximize_btn = TitleBarButton("□")
        self.maximize_btn.clicked.connect(self.toggle_maximize)
        layout.addWidget(self.maximize_btn)
        
        # 关闭按钮
        self.close_btn = TitleBarButton("×")
        self.close_btn.setStyleSheet(self.close_btn.styleSheet() + """
            QPushButton:hover {
                background-color: #C83C3C;
                color: white;
            }
        """)
        self.close_btn.clicked.connect(self.close_window)
        layout.addWidget(self.close_btn)
        
        self.setLayout(layout)
        self.setFixedHeight(30)
    
    def minimize_window(self):
        # 修改这里，调用父窗口的最小化到托盘方法
        self.parent.minimize_to_tray()
    
    def toggle_maximize(self):
        if self.parent.isMaximized():
            self.parent.showNormal()
            self.maximize_btn.setText("□")
        else:
            self.parent.showMaximized()
            self.maximize_btn.setText("❐")
    
    def close_window(self):
        # 修复：点击X按钮时真正关闭程序
        self.parent.close_application()
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.parent.frameGeometry().topLeft()
            event.accept()
    
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and hasattr(self, 'drag_position'):
            self.parent.move(event.globalPos() - self.drag_position)
            event.accept()

# 启动工作线程
class LaunchThread(QThread):
    finished = pyqtSignal()
    status_update = pyqtSignal(str, bool, str)  # path, running, process_name
    
    def __init__(self, programs, parent=None):
        super().__init__(parent)
        self.programs = programs
        self.is_running = True
    
    def run(self):
        # 先启动普通exe程序
        for row in self.programs:
            if not self.is_running:
                break
            path = row.get_program_path()
            if row.is_valid() and not row.is_uwp:
                try:
                    process_name = os.path.basename(path)
                    self.status_update.emit(path, True, process_name)
                    # 以管理员权限启动程序
                    ctypes.windll.shell32.ShellExecuteW(
                        None, "runas", path, None, None, 1
                    )
                    # 稍等让程序启动
                    time.sleep(0.5)
                except Exception as e:
                    print(f"启动程序出错: {e}")
        
        # 再启动UWP程序
        for row in self.programs:
            if not self.is_running:
                break
            path = row.get_program_path()
            if row.is_valid() and row.is_uwp:
                try:
                    self.status_update.emit(path, True, row.process_name or "UWP应用")
                    # 启动UWP应用
                    if os.path.exists(path):
                        os.startfile(path)
                    # 稍等让程序启动
                    time.sleep(0.5)
                except Exception as e:
                    print(f"启动UWP应用出错: {e}")
        
        self.finished.emit()
    
    def stop(self):
        self.is_running = False

# 关闭工作线程
class CloseThread(QThread):
    finished = pyqtSignal()
    status_update = pyqtSignal(str, bool)  # path, running
    
    def __init__(self, programs, parent=None):
        super().__init__(parent)
        self.programs = programs
        self.is_running = True
    
    def run(self):
        # 关闭所有程序
        for row in self.programs:
            if not row.is_valid():
                continue
            
            try:
                # 获取进程名称
                process_name = row.selected_process or row.process_name or os.path.basename(row.get_program_path())
                
                # 关闭进程及其子进程
                if psutil:
                    for proc in psutil.process_iter(['pid', 'name', 'exe']):
                        try:
                            info = proc.info
                            proc_name = info['name'].lower()
                            # 检查进程名是否匹配
                            if process_name.lower() in proc_name:
                                # 结束进程树
                                parent = psutil.Process(info['pid'])
                                children = parent.children(recursive=True)
                                
                                # 先结束子进程
                                for child in children:
                                    try:
                                        child.terminate()
                                    except:
                                        pass
                                
                                # 等待子进程结束
                                psutil.wait_procs(children, timeout=3)
                                
                                # 结束父进程
                                try:
                                    parent.terminate()
                                    parent.wait(3)
                                except:
                                    try:
                                        parent.kill()
                                    except:
                                        pass
                                
                                self.status_update.emit(row.get_program_path(), False)
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                            continue
            except Exception as e:
                print(f"关闭程序出错: {e}")
        
        self.finished.emit()
    
    def stop(self):
        self.is_running = False

# 主窗口
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.program_rows = []
        self.config_file = "launcher_config.json"
        self.tray_icon = None
        self.launch_thread = None
        self.close_thread = None
        self.is_closing = False  # 标记是否正在关闭程序
        
        # 设置应用图标 - 修复图标显示问题
        app_icon = get_app_icon()
        self.setWindowIcon(app_icon)
        
        # 设置窗口属性
        self.setWindowTitle("程序启动管理器")
        # 设置无边框窗口和自定义标题栏
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # 创建主部件
        central_widget = QWidget()
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #2B2B2B;
                color: #A9B7C6;
            }
        """)
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 添加自定义标题栏
        self.title_bar = CustomTitleBar(self)
        main_layout.addWidget(self.title_bar)
        
        # 内容区域
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 15, 20, 15)
        content_layout.setSpacing(10)
        
        # 顶部按钮区域
        top_buttons_layout = QHBoxLayout()
        self.launch_all_btn = QPushButton("一键开启")
        self.launch_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #4C7A43;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5A9050;
            }
            QPushButton:pressed {
                background-color: #3E6436;
            }
        """)
        self.launch_all_btn.clicked.connect(self.launch_all_programs)
        top_buttons_layout.addWidget(self.launch_all_btn)
        
        self.close_all_btn = QPushButton("一键关闭")
        self.close_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #A04040;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #B54A4A;
            }
            QPushButton:pressed {
                background-color: #863636;
            }
        """)
        self.close_all_btn.clicked.connect(self.close_all_programs)
        top_buttons_layout.addWidget(self.close_all_btn)
        
        self.add_program_btn = QPushButton("添加程序")
        self.add_program_btn.setStyleSheet("""
            QPushButton {
                background-color: #3C6496;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #4A7BB0;
            }
            QPushButton:pressed {
                background-color: #32527A;
            }
        """)
        self.add_program_btn.clicked.connect(self.add_program_row)
        top_buttons_layout.addWidget(self.add_program_btn)
        
        self.save_config_btn = QPushButton("保存配置")
        self.save_config_btn.setStyleSheet("""
            QPushButton {
                background-color: #6C6C6C;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7C7C7C;
            }
            QPushButton:pressed {
                background-color: #5C5C5C;
            }
        """)
        self.save_config_btn.clicked.connect(self.save_config)
        top_buttons_layout.addWidget(self.save_config_btn)
        
        top_buttons_layout.addStretch()
        content_layout.addLayout(top_buttons_layout)
        
        # 滚动区域用于程序列表
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: 1px solid #4B4B4B;
                background: #373737;
                width: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #555555;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)
        
        self.scroll_content = QWidget()
        self.scroll_content.setStyleSheet("background-color: transparent;")
        self.programs_layout = QVBoxLayout(self.scroll_content)
        self.programs_layout.setContentsMargins(0, 0, 0, 0)
        self.programs_layout.setSpacing(8)
        
        # 添加程序行标题
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("程序路径"), 4)
        header_layout.addWidget(QLabel("操作"), 1)
        header_layout.addWidget(QLabel("状态"), 2)
        header_layout.addWidget(QLabel("进程管理"), 1)
        header_layout.addWidget(QLabel(""), 1)
        
        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        header_widget.setStyleSheet("""
            QLabel {
                font-weight: bold;
                color: #a9b7c6;
                min-height: 25px;
            }
        """)
        self.programs_layout.addWidget(header_widget)
        
        self.scroll_area.setWidget(self.scroll_content)
        content_layout.addWidget(self.scroll_area, 1)
        
        main_layout.addWidget(content_widget, 1)
        
        # 加载配置
        self.load_config()
        
        # 如果没有程序行，添加3个默认行
        if not self.program_rows:
            for _ in range(3):
                self.add_program_row()
        
        # 设置系统托盘
        self.setup_system_tray()
    
    def setup_system_tray(self):
        # 创建系统托盘图标 - 使用应用图标
        self.tray_icon = QSystemTrayIcon(get_app_icon(), self)
        
        # 创建托盘菜单
        tray_menu = QMenu()
        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)
        
        launch_action = QAction("启动所有程序", self)
        launch_action.triggered.connect(self.launch_all_programs)
        tray_menu.addAction(launch_action)
        
        close_action = QAction("关闭所有程序", self)
        close_action.triggered.connect(self.close_all_programs)
        tray_menu.addAction(close_action)
        
        tray_menu.addSeparator()
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close_application)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()
        self.tray_icon.setToolTip("程序启动管理器")
    
    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()
    
    def show_window(self):
        self.show()
        self.raise_()
        if self.isMinimized():
            self.showNormal()
        self.activateWindow()
    
    # 新增方法：最小化到系统托盘
    def minimize_to_tray(self):
        """最小化到系统托盘"""
        self.hide()
        self.tray_icon.showMessage(
            "程序启动管理器",
            "程序已最小化到系统托盘",
            QSystemTrayIcon.Information,
            2000
        )
    
    def close_application(self):
        """真正关闭应用程序"""
        self.is_closing = True
        # 关闭所有程序
        self.close_all_programs()
        
        # 确保线程完成
        if self.launch_thread and self.launch_thread.isRunning():
            self.launch_thread.stop()
            self.launch_thread.wait()
        
        if self.close_thread and self.close_thread.isRunning():
            self.close_thread.stop()
            self.close_thread.wait()
        
        # 退出系统托盘
        if self.tray_icon:
            self.tray_icon.hide()
        
        # 退出应用
        QApplication.quit()
    
    def closeEvent(self, event):
        if self.is_closing:
            # 正在关闭程序，允许事件通过
            event.accept()
            return
        
        # 拦截关闭事件，改为最小化到托盘
        event.ignore()
        self.hide()
        
        # 仅在窗口可见时显示提示
        if self.isVisible():
            self.tray_icon.showMessage(
                "程序启动管理器",
                "程序已最小化到系统托盘",
                QSystemTrayIcon.Information,
                2000
            )
    
    def add_program_row(self):
        row = ProgramRow(manager=self)
        self.program_rows.append(row)
        self.programs_layout.addWidget(row)
        
        # 如果行数超过5，启用滚动条
        if len(self.program_rows) > 5:
            self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    
    def remove_program_row(self, row):
        if row in self.program_rows:
            self.program_rows.remove(row)
            row.deleteLater()
    
    def launch_all_programs(self):
        if self.launch_thread and self.launch_thread.isRunning():
            QMessageBox.warning(self, "警告", "程序启动中，请稍候...")
            return
        
        valid_rows = [row for row in self.program_rows if row.is_valid()]
        if not valid_rows:
            QMessageBox.warning(self, "警告", "没有有效的程序路径")
            return
        
        # 重置状态
        for row in valid_rows:
            row.set_status(False)
        
        # 创建并启动线程
        self.launch_thread = LaunchThread(valid_rows)
        self.launch_thread.status_update.connect(self.update_program_status)
        self.launch_thread.finished.connect(self.on_launch_finished)
        self.launch_thread.start()
        
        self.launch_all_btn.setEnabled(False)
        self.close_all_btn.setEnabled(False)
        self.add_program_btn.setEnabled(False)
        self.save_config_btn.setEnabled(False)
    
    def update_program_status(self, path, running, process_name):
        for row in self.program_rows:
            if row.get_program_path() == path:
                row.set_status(running, process_name)
                if running and not row.process_name:
                    row.process_name = process_name
                break
    
    def on_launch_finished(self):
        self.launch_all_btn.setEnabled(True)
        self.close_all_btn.setEnabled(True)
        self.add_program_btn.setEnabled(True)
        self.save_config_btn.setEnabled(True)
    
    def close_all_programs(self):
        if self.close_thread and self.close_thread.isRunning():
            QMessageBox.warning(self, "警告", "程序关闭中，请稍候...")
            return
        
        valid_rows = [row for row in self.program_rows if row.is_valid()]
        if not valid_rows:
            QMessageBox.warning(self, "警告", "没有有效的程序路径")
            return
        
        # 创建并启动线程
        self.close_thread = CloseThread(valid_rows)
        self.close_thread.status_update.connect(self.update_close_status)
        self.close_thread.finished.connect(self.on_close_finished)
        self.close_thread.start()
        
        self.launch_all_btn.setEnabled(False)
        self.close_all_btn.setEnabled(False)
        self.add_program_btn.setEnabled(False)
        self.save_config_btn.setEnabled(False)
    
    def update_close_status(self, path, running):
        for row in self.program_rows:
            if row.get_program_path() == path:
                row.set_status(False)
                break
    
    def on_close_finished(self):
        self.launch_all_btn.setEnabled(True)
        self.close_all_btn.setEnabled(True)
        self.add_program_btn.setEnabled(True)
        self.save_config_btn.setEnabled(True)
    
    def save_config(self):
        config = []
        for row in self.program_rows:
            path = row.get_program_path()
            if path and os.path.exists(path):
                config.append({
                    "path": path,
                    "is_uwp": row.is_uwp,
                    "process_name": row.process_name,
                    "selected_process": row.selected_process
                })
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, "成功", "配置已保存")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存配置失败: {str(e)}")
    
    def load_config(self):
        if not os.path.exists(self.config_file):
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 清空现有行
            for row in self.program_rows[:]:
                self.remove_program_row(row)
            
            # 加载配置
            for item in config:
                row = ProgramRow(manager=self)
                row.path_input.setText(item.get("path", ""))
                row.is_uwp = item.get("is_uwp", False)
                row.process_name = item.get("process_name")
                row.selected_process = item.get("selected_process")
                self.program_rows.append(row)
                self.programs_layout.addWidget(row)
            
            # 至少保留3行
            while len(self.program_rows) < 3:
                self.add_program_row()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载配置失败: {str(e)}")

# 检查依赖
def check_dependencies():
    missing_deps = []
    if not psutil:
        missing_deps.append("psutil (用于进程管理)")
    if not Dispatch:
        missing_deps.append("pywin32 (用于UWP应用管理)")
    
    if missing_deps:
        msg = "缺少以下依赖库，部分功能将受限:\n" + "\n".join(missing_deps)
        msg += "\n\n建议使用以下命令安装:\n"
        msg += "pip install psutil pywin32"
        app = QApplication(sys.argv)
        QMessageBox.warning(None, "依赖缺失", msg)
        app.quit()

def main():
    # 检查管理员权限
    if not is_admin():
        if not run_as_admin():
            app = QApplication(sys.argv)
            QMessageBox.critical(None, "权限错误", "需要管理员权限才能运行此程序")
            sys.exit(1)
        else:
            sys.exit(0)
    
    # 检查依赖
    check_dependencies()
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # 设置应用字体
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    
    # 设置全局应用图标 - 修复任务栏图标
    app_icon = get_app_icon()
    app.setWindowIcon(app_icon)
    
    # 创建主窗口
    window = MainWindow()
    window.resize(900, 600)
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()