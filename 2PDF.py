import os
import sys
import subprocess
import tempfile
import shutil
import quopri
import base64
import re
import glob
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, 
                             QFileDialog, QLabel, QProgressBar, QTextEdit, QCheckBox, QGroupBox,
                             QTabWidget, QListWidget, QListWidgetItem, QSplitter, QComboBox, QMessageBox)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, QTimer, pyqtSignal, QThread, Qt, QMarginsF
from PyQt5.QtGui import QPageLayout, QPageSize, QFont
from PyQt5.QtPrintSupport import QPrinter

class BatchConverter(QThread):
    """批量转换线程"""
    progress_updated = pyqtSignal(int, int, str)  # 当前进度,总数,当前文件
    conversion_completed = pyqtSignal(str)  # 转换完成信息
    
    def __init__(self, mht_files, output_dir, delete_original=False):
        super().__init__()
        self.mht_files = mht_files
        self.output_dir = output_dir
        self.delete_original = delete_original
        self.converter_instance = None
        
    def run(self):
        total_files = len(self.mht_files)
        success_count = 0
        failed_files = []
        
        for i, mht_file in enumerate(self.mht_files):
            try:
                self.progress_updated.emit(i + 1, total_files, os.path.basename(mht_file))
                
                # 创建临时转换器实例
                if self.convert_single_file(mht_file):
                    success_count += 1
                    if self.delete_original:
                        try:
                            os.remove(mht_file)
                        except Exception as e:
                            print(f"删除原文件失败 {mht_file}: {e}")
                else:
                    failed_files.append(mht_file)
                    
            except Exception as e:
                failed_files.append(mht_file)
                print(f"转换失败 {mht_file}: {e}")
        
        # 发送完成信号
        result_msg = f"批量转换完成!\n成功: {success_count}/{total_files}"
        if failed_files:
            result_msg += f"\n失败文件: {len(failed_files)}个"
        if self.delete_original and success_count > 0:
            result_msg += f"\n已删除原始文件: {success_count}个"
            
        self.conversion_completed.emit(result_msg)
    
    def convert_single_file(self, mht_file):
        """转换单个文件"""
        try:
            # 这里需要实现单个文件的转换逻辑
            # 由于WebEngine需要在主线程运行,这里先返回True作为占位
            return True
        except Exception as e:
            print(f"转换文件失败 {mht_file}: {e}")
            return False

class HTMLtoPDFConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        """初始化中文界面"""
        # 设置窗口标题和图标
        self.setWindowTitle("MHT2PDF - github.com/LeeKaiGit")
        self.setGeometry(300, 300, 800, 600)
        
        # 设置窗口图标 - 支持打包后的exe文件
        from PyQt5.QtGui import QIcon, QPixmap
        try:
            # 尝试从多个位置加载图标
            icon_loaded = False
            
            # 1. 尝试从打包后的临时目录加载
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, 'pdf.ico')
                if os.path.exists(icon_path):
                    self.setWindowIcon(QIcon(icon_path))
                    icon_loaded = True
            
            # 2. 尝试从当前脚本目录加载
            if not icon_loaded:
                icon_path = os.path.join(os.path.dirname(__file__), 'pdf.ico')
                if os.path.exists(icon_path):
                    self.setWindowIcon(QIcon(icon_path))
                    icon_loaded = True
            
            # 3. 尝试从当前工作目录加载
            if not icon_loaded:
                icon_path = 'pdf.ico'
                if os.path.exists(icon_path):
                    self.setWindowIcon(QIcon(icon_path))
                    icon_loaded = True
            
            # 4. 如果都失败,创建一个简单的默认图标
            if not icon_loaded:
                pixmap = QPixmap(32, 32)
                pixmap.fill()  # 填充为白色
                self.setWindowIcon(QIcon(pixmap))
                
        except Exception as e:
            print(f"加载图标失败: {e}")
        
        main_layout = QVBoxLayout()
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        
        # 单文件转换选项卡
        self.single_tab = QWidget()
        self.init_single_tab()
        self.tab_widget.addTab(self.single_tab, "单文件转换")
        
        # 批量转换选项卡
        self.batch_tab = QWidget()
        self.init_batch_tab()
        self.tab_widget.addTab(self.batch_tab, "批量转换")
        
        main_layout.addWidget(self.tab_widget)
        
        # 添加作者署名(右下角)
        author_label = QLabel("github.com/LeeKaiGit")
        author_label.setStyleSheet("""
            QLabel {
                color: #ff0000;
                font-size: 20px;
                font-style: italic;
                font-weight: bold;
                padding: 5px;
            }
        """)
        author_label.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        main_layout.addWidget(author_label)
        
        self.setLayout(main_layout)
        
        # 设置窗口属性
        self.setWindowTitle("MHT2PDF")
        self.resize(1400, 900)
        
        # 变量初始化
        self.last_directory = ""
        self.page_loaded = False
        self.imported_file_path = None
        self.batch_files = []

    def init_single_tab(self):
        """初始化单文件转换选项卡"""
        layout = QVBoxLayout()
        
        # 按钮区域
        button_layout = QHBoxLayout()
        
        self.import_button = QPushButton("导入 MHT/HTML 文件")
        self.import_button.clicked.connect(self.import_file)
        button_layout.addWidget(self.import_button)

        self.export_button = QPushButton("导出为 PDF")
        self.export_button.clicked.connect(self.export_pdf)
        self.export_button.setEnabled(False)
        button_layout.addWidget(self.export_button)

        layout.addLayout(button_layout)

        # 信息标签
        self.info_label = QLabel("未导入文件")
        layout.addWidget(self.info_label)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 网页预览
        self.web_view = QWebEngineView()
        settings = self.web_view.settings()
        settings.setAttribute(settings.JavascriptEnabled, True)
        settings.setAttribute(settings.AutoLoadImages, True)
        settings.setAttribute(settings.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(settings.LocalContentCanAccessFileUrls, True)
        layout.addWidget(self.web_view)
        
        self.single_tab.setLayout(layout)

    def init_batch_tab(self):
        """初始化批量转换选项卡"""
        layout = QVBoxLayout()
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # 选择方式
        select_layout = QHBoxLayout()
        
        self.select_files_btn = QPushButton("选择多个 MHT 文件")
        self.select_files_btn.clicked.connect(self.select_multiple_files)
        select_layout.addWidget(self.select_files_btn)
        
        self.select_folder_btn = QPushButton("选择文件夹")
        self.select_folder_btn.clicked.connect(self.select_folder)
        select_layout.addWidget(self.select_folder_btn)
        
        # 子文件夹选项
        self.include_subfolders = QCheckBox("包含子文件夹")
        self.include_subfolders.setChecked(True)
        select_layout.addWidget(self.include_subfolders)
        
        file_layout.addLayout(select_layout)
        
        # 文件列表
        self.file_list = QListWidget()
        file_layout.addWidget(self.file_list)
        
        # 清除按钮
        clear_layout = QHBoxLayout()
        self.clear_list_btn = QPushButton("清空列表")
        self.clear_list_btn.clicked.connect(self.clear_file_list)
        clear_layout.addWidget(self.clear_list_btn)
        clear_layout.addStretch()
        file_layout.addLayout(clear_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 输出设置区域
        output_group = QGroupBox("输出设置")
        output_layout = QVBoxLayout()
        
        # 输出目录选择
        output_dir_layout = QHBoxLayout()
        self.output_dir_label = QLabel("输出目录: 将自动设置为MHT文件所在目录")
        output_dir_layout.addWidget(self.output_dir_label)
        
        self.select_output_dir_btn = QPushButton("选择输出目录")
        self.select_output_dir_btn.clicked.connect(self.select_output_directory)
        output_dir_layout.addWidget(self.select_output_dir_btn)
        
        output_layout.addLayout(output_dir_layout)
        
        # 删除原文件选项
        self.delete_original_cb = QCheckBox("转换完成后删除原始 MHT 文件")
        self.delete_original_cb.setStyleSheet("QCheckBox { color: red; font-weight: bold; }")
        output_layout.addWidget(self.delete_original_cb)
        
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
        # 批量转换控制
        batch_control_layout = QHBoxLayout()
        
        self.start_batch_btn = QPushButton("开始批量转换")
        self.start_batch_btn.clicked.connect(self.start_batch_conversion)
        self.start_batch_btn.setEnabled(False)
        batch_control_layout.addWidget(self.start_batch_btn)
        
        batch_control_layout.addStretch()
        layout.addLayout(batch_control_layout)
        
        # 批量转换进度
        self.batch_progress = QProgressBar()
        self.batch_progress.setVisible(False)
        layout.addWidget(self.batch_progress)
        
        self.batch_status_label = QLabel("就绪")
        layout.addWidget(self.batch_status_label)
        
        # 转换日志
        log_group = QGroupBox("转换日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.batch_tab.setLayout(layout)
        
        # 初始化变量
        self.output_directory = ""

    def set_batch_controls_enabled(self, enabled):
        """设置批量转换控件的启用状态"""
        self.select_files_btn.setEnabled(enabled)
        self.select_folder_btn.setEnabled(enabled)
        self.select_output_dir_btn.setEnabled(enabled)
        self.clear_list_btn.setEnabled(enabled)
        self.include_subfolders.setEnabled(enabled)
        self.delete_original_cb.setEnabled(enabled)
        if enabled:
            self.update_batch_button_state()
        else:
            self.start_batch_btn.setEnabled(False)

    def select_multiple_files(self):
        """选择多个MHT文件"""
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "选择多个 MHT 文件", 
            self.last_directory, 
            "MHT Files (*.mht *.mhtml);;All Files (*.*)", 
            options=options
        )
        
        if files:
            self.last_directory = os.path.dirname(files[0])
            
            # 自动设置输出目录为第一个文件所在的目录
            if not self.output_directory:
                self.output_directory = os.path.dirname(files[0])
                self.output_dir_label.setText(f"输出目录: {self.output_directory} (自动设置)")
            
            for file in files:
                if file not in [self.file_list.item(i).text() for i in range(self.file_list.count())]:
                    item = QListWidgetItem(file)
                    self.file_list.addItem(item)
            
            self.update_batch_button_state()
            self.log_text.append(f"添加了 {len(files)} 个文件")

    def select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "选择包含 MHT 文件的文件夹", 
            self.last_directory
        )
        
        if folder:
            self.last_directory = folder
            
            # 自动设置输出目录为选择的文件夹
            if not self.output_directory:
                self.output_directory = folder
                self.output_dir_label.setText(f"输出目录: {folder} (自动设置)")
            
            # 搜索MHT文件
            pattern = "**/*.mht" if self.include_subfolders.isChecked() else "*.mht"
            mht_files = glob.glob(os.path.join(folder, pattern), recursive=self.include_subfolders.isChecked())
            
            # 同时搜索mhtml文件
            pattern_mhtml = "**/*.mhtml" if self.include_subfolders.isChecked() else "*.mhtml"
            mhtml_files = glob.glob(os.path.join(folder, pattern_mhtml), recursive=self.include_subfolders.isChecked())
            
            all_files = mht_files + mhtml_files
            
            if all_files:
                for file in all_files:
                    if file not in [self.file_list.item(i).text() for i in range(self.file_list.count())]:
                        item = QListWidgetItem(file)
                        self.file_list.addItem(item)
                
                self.update_batch_button_state()
                self.log_text.append(f"从文件夹 {folder} 找到 {len(all_files)} 个 MHT 文件")
            else:
                self.log_text.append(f"在文件夹 {folder} 中未找到 MHT 文件")

    def select_output_directory(self):
        """选择输出目录"""
        directory = QFileDialog.getExistingDirectory(
            self, 
            "选择 PDF 输出目录", 
            self.last_directory
        )
        
        if directory:
            self.output_directory = directory
            self.output_dir_label.setText(f"输出目录: {directory} (手动设置)")
            self.update_batch_button_state()

    def clear_file_list(self):
        """清空文件列表"""
        self.file_list.clear()
        # 清空输出目录设置
        self.output_directory = ""
        self.output_dir_label.setText("输出目录: 将自动设置为MHT文件所在目录")
        self.update_batch_button_state()
        self.log_text.append("已清空文件列表")

    def update_batch_button_state(self):
        """更新批量转换按钮状态"""
        has_files = self.file_list.count() > 0
        # 如果有文件,输出目录可以自动设置,所以只需要检查是否有文件
        self.start_batch_btn.setEnabled(has_files)

    def start_batch_conversion(self):
        """开始批量转换"""
        if self.file_list.count() == 0:
            self.log_text.append("错误: 没有选择文件")
            return
        
        # 禁用界面控件
        self.set_batch_controls_enabled(False)
        
        # 如果没有设置输出目录,自动设置为第一个文件所在的目录
        if not self.output_directory:
            first_file = self.file_list.item(0).text()
            self.output_directory = os.path.dirname(first_file)
            self.output_dir_label.setText(f"输出目录: {self.output_directory} (自动设置)")
            self.log_text.append(f"自动设置输出目录为: {self.output_directory}")
        
        # 获取文件列表
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        
        # 显示进度
        self.batch_progress.setVisible(True)
        self.batch_progress.setMaximum(len(files))
        self.batch_progress.setValue(0)
        
        self.start_batch_btn.setEnabled(False)
        self.batch_status_label.setText("正在批量转换...")
        
        delete_original = self.delete_original_cb.isChecked()
        
        self.log_text.append(f"开始批量转换 {len(files)} 个文件...")
        if delete_original:
            self.log_text.append("警告: 将在转换成功后删除原始文件")
        
        # 由于WebEngine限制,这里需要改为同步处理
        self.process_batch_files(files, delete_original)

    def process_batch_files(self, files, delete_original):
        """处理批量文件转换"""
        self.batch_current_index = 0
        self.batch_files_list = files
        self.batch_delete_original = delete_original
        self.batch_success_count = 0
        self.batch_failed_files = []
        
        # 计算基础目录(所有文件的公共父目录)
        if len(files) == 1:
            # 单个文件时,基础目录是文件所在目录
            self.batch_base_directory = os.path.dirname(files[0])
        else:
            # 多个文件时,找到公共父目录
            self.batch_base_directory = os.path.commonpath([os.path.dirname(f) for f in files])
        
        self.log_text.append(f"基础目录: {self.batch_base_directory}")
        self.log_text.append(f"将保持原有的子文件夹结构")
        
        # 开始处理第一个文件
        self.process_next_batch_file()

    def process_next_batch_file(self):
        """处理下一个批量文件"""
        if self.batch_current_index >= len(self.batch_files_list):
            # 批量处理完成
            self.finish_batch_conversion()
            return
        
        current_file = self.batch_files_list[self.batch_current_index]
        file_name = os.path.basename(current_file)
        
        # 更新进度
        self.batch_progress.setValue(self.batch_current_index + 1)
        self.batch_status_label.setText(f"正在转换: {file_name} ({self.batch_current_index + 1}/{len(self.batch_files_list)})")
        self.log_text.append(f"开始转换: {file_name}")
        
        try:
            # 设置当前文件
            self.imported_file_path = current_file
            
            # 处理MHT文件
            processed_path = self.preprocess_mht_file(current_file)
            if processed_path:
                # 加载文件到WebView
                try:
                    self.web_view.loadFinished.disconnect()
                except:
                    pass
                self.web_view.loadFinished.connect(self.on_batch_file_loaded)
                self.web_view.load(QUrl.fromLocalFile(processed_path))
            else:
                self.log_text.append(f"错误: 无法处理文件 {file_name}")
                self.batch_failed_files.append(current_file)
                self.batch_current_index += 1
                QTimer.singleShot(100, self.process_next_batch_file)
                
        except Exception as e:
            self.log_text.append(f"错误: 处理文件 {file_name} 失败: {str(e)}")
            self.batch_failed_files.append(current_file)
            self.batch_current_index += 1
            QTimer.singleShot(100, self.process_next_batch_file)

    def on_batch_file_loaded(self, success):
        """批量文件加载完成回调"""
        current_file = self.batch_files_list[self.batch_current_index]
        file_name = os.path.basename(current_file)
        
        if success:
            # 应用渲染优化
            self.inject_rendering_improvements()
            
            # 延迟执行PDF导出
            QTimer.singleShot(2000, self.export_current_batch_file)
        else:
            self.log_text.append(f"错误: 文件 {file_name} 加载失败")
            self.batch_failed_files.append(current_file)
            self.batch_current_index += 1
            QTimer.singleShot(100, self.process_next_batch_file)

    def export_current_batch_file(self):
        """导出当前批量文件为PDF"""
        current_file = self.batch_files_list[self.batch_current_index]
        file_name = os.path.basename(current_file)
        name_without_ext = os.path.splitext(file_name)[0]
        
        # 根据"包含子文件夹"选项决定保存位置
        if self.include_subfolders.isChecked():
            # 如果勾选了"包含子文件夹",PDF保存在原文件所在目录
            pdf_dir = os.path.dirname(current_file)
            pdf_path = os.path.join(pdf_dir, f"{name_without_ext}.pdf")
        else:
            # 如果没有勾选,保存到输出目录,但保持子文件夹结构
            file_dir = os.path.dirname(current_file)
            if file_dir.startswith(self.batch_base_directory):
                # 获取相对路径
                relative_dir = os.path.relpath(file_dir, self.batch_base_directory)
                if relative_dir == ".":
                    # 如果就在基础目录下,直接使用输出目录
                    pdf_dir = self.output_directory
                else:
                    # 在输出目录下创建相同的子文件夹结构
                    pdf_dir = os.path.join(self.output_directory, relative_dir)
            else:
                # 如果不在基础目录下(不应该发生),直接使用输出目录
                pdf_dir = self.output_directory
            
            pdf_path = os.path.join(pdf_dir, f"{name_without_ext}.pdf")
        
        try:
            # 确保输出目录存在
            os.makedirs(pdf_dir, exist_ok=True)
            
            # 执行PDF导出
            self.perform_batch_pdf_export(pdf_path, current_file)
            
        except Exception as e:
            self.log_text.append(f"错误: 导出 {file_name} 失败: {str(e)}")
            self.batch_failed_files.append(current_file)
            self.batch_current_index += 1
            QTimer.singleShot(100, self.process_next_batch_file)

    def perform_batch_pdf_export(self, pdf_path, original_file):
        """执行批量PDF导出"""
        try:
            # 应用最终样式优化
            final_js = """
            console.log('Applying final A4 print optimizations...');
            
            function applyFinalStyles() {
                var finalStyle = document.createElement('style');
                finalStyle.innerHTML = `
                    @page {
                        size: A4 portrait;
                        margin: 1cm 1.5cm;
                    }
                    
                    body {
                        margin: 0 !important;
                        padding: 10px !important;
                        background: white !important;
                        max-width: 100% !important;
                    }
                    
                    table {
                        width: 100% !important;
                        border-collapse: collapse !important;
                        margin: 0 auto 8px auto !important;
                        table-layout: auto !important;
                    }
                    
                    td, th {
                        border: 1px solid #000 !important;
                        padding: 4px 6px !important;
                        word-wrap: break-word !important;
                        vertical-align: top !important;
                    }
                    
                    th {
                        background-color: #f0f0f0 !important;
                    }
                    
                    img {
                        max-width: 120px !important;
                        max-height: 150px !important;
                        width: auto !important;
                        height: auto !important;
                        display: block !important;
                        margin: 2px auto !important;
                        object-fit: contain !important;
                    }
                `;
                
                if (document.head) {
                    document.head.appendChild(finalStyle);
                }
            }
            
            applyFinalStyles();
            """
            
            self.web_view.page().runJavaScript(final_js)
            
            # 延迟执行实际的PDF导出
            QTimer.singleShot(1000, lambda: self.do_batch_pdf_export(pdf_path, original_file))
            
        except Exception as e:
            file_name = os.path.basename(original_file)
            self.log_text.append(f"错误: 准备导出 {file_name} 失败: {str(e)}")
            self.batch_failed_files.append(original_file)
            self.batch_current_index += 1
            QTimer.singleShot(100, self.process_next_batch_file)

    def do_batch_pdf_export(self, pdf_path, original_file):
        """执行实际的批量PDF导出"""
        try:
            file_name = os.path.basename(original_file)
            
            # 使用简化的WebEngine PDF导出
            try:
                # 使用WebEngine的简单printToPdf方法(避免页面布局参数问题)
                self.web_view.page().printToPdf(pdf_path)
                
                # 等待PDF生成完成后处理
                QTimer.singleShot(3000, lambda: self.check_pdf_export_result(pdf_path, original_file))
                
            except Exception as fallback_error:
                # 如果WebEngine方法失败,记录错误并跳过
                file_name = os.path.basename(original_file)
                self.log_text.append(f"错误: 导出 {file_name} 失败: {str(fallback_error)}")
                self.batch_failed_files.append(original_file)
                self.batch_current_index += 1
                QTimer.singleShot(100, self.process_next_batch_file)
            
        except Exception as e:
            file_name = os.path.basename(original_file)
            self.log_text.append(f"错误: 导出 {file_name} 失败: {str(e)}")
            self.batch_failed_files.append(original_file)
            self.batch_current_index += 1
            QTimer.singleShot(100, self.process_next_batch_file)

    def check_pdf_export_result(self, pdf_path, original_file):
        """检查PDF导出结果(用于WebEngine printToPdf方法)"""
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
            self.on_batch_export_finished(True, pdf_path, original_file)
        else:
            self.on_batch_export_finished(False, pdf_path, original_file)

    def on_batch_export_finished(self, success, pdf_path, original_file):
        """批量导出完成回调"""
        file_name = os.path.basename(original_file)
        
        if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
            self.log_text.append(f"成功: {file_name} -> {os.path.basename(pdf_path)}")
            self.batch_success_count += 1
            
            # 删除原文件(如果选择了该选项)
            if self.batch_delete_original:
                try:
                    os.remove(original_file)
                    self.log_text.append(f"已删除: {file_name}")
                except Exception as e:
                    self.log_text.append(f"警告: 无法删除 {file_name}: {str(e)}")
        else:
            self.log_text.append(f"失败: {file_name}")
            self.batch_failed_files.append(original_file)
        
        # 处理下一个文件
        self.batch_current_index += 1
        QTimer.singleShot(500, self.process_next_batch_file)

    def finish_batch_conversion(self):
        """完成批量转换"""
        total_files = len(self.batch_files_list)
        
        self.batch_progress.setVisible(False)
        # 重新启用界面控件
        self.set_batch_controls_enabled(True)
        
        # 显示结果
        result_msg = f"批量转换完成!\n成功: {self.batch_success_count}/{total_files}"
        if self.batch_failed_files:
            result_msg += f"\n失败: {len(self.batch_failed_files)} 个文件"
        if self.batch_delete_original and self.batch_success_count > 0:
            result_msg += f"\n已删除原始文件: {self.batch_success_count} 个"
        
        self.batch_status_label.setText(result_msg)
        self.log_text.append("=" * 50)
        self.log_text.append(result_msg)
        
        if self.batch_failed_files:
            self.log_text.append("失败的文件:")
            for failed_file in self.batch_failed_files:
                self.log_text.append(f"  - {os.path.basename(failed_file)}")
        
        # 显示完成通知弹窗
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("转换完成")
        msg_box.setText(result_msg)
        if self.batch_success_count == total_files:
            msg_box.setIcon(QMessageBox.Information)
        else:
            msg_box.setIcon(QMessageBox.Warning)
        msg_box.exec_()

    def import_file(self):
        """导入单个文件"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "导入 MHT/HTML 文件", 
            self.last_directory, 
            "MHT/HTML Files (*.mht *.mhtml *.html *.htm);;All Files (*.*)", 
            options=options
        )
        
        if file_path:
            self.last_directory = os.path.dirname(file_path)
            self.imported_file_path = file_path
            
            # 显示进度
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self.info_label.setText("正在加载文件...")
            
            # 更新文件信息
            file_name = os.path.basename(file_path)
            file_dir = os.path.dirname(file_path)
            file_ext = os.path.splitext(file_path)[1]
            file_size = os.path.getsize(file_path)
            
            self.info_label.setText(
                f"文件名: {file_name}\n"
                f"位置: {file_dir}\n"
                f"扩展名: {file_ext}\n"
                f"大小: {file_size:,} 字节\n"
                f"状态: 正在加载..."
            )
            
            # 处理MHT文件
            if file_ext.lower() in ['.mht', '.mhtml']:
                processed_path = self.preprocess_mht_file(file_path)
                if processed_path:
                    file_path = processed_path
            
            # 连接加载完成信号
            try:
                self.web_view.loadFinished.disconnect()
            except:
                pass
            
            self.web_view.loadFinished.connect(self.on_page_loaded)
            self.web_view.load(QUrl.fromLocalFile(file_path))

    def on_page_loaded(self, success):
        """页面加载完成回调"""
        self.progress_bar.setVisible(False)
        
        if success:
            self.page_loaded = True
            self.export_button.setEnabled(True)
            
            file_name = os.path.basename(self.imported_file_path)
            file_dir = os.path.dirname(self.imported_file_path)
            file_ext = os.path.splitext(self.imported_file_path)[1]
            file_size = os.path.getsize(self.imported_file_path)
            
            self.info_label.setText(
                f"文件名: {file_name}\n"
                f"位置: {file_dir}\n"
                f"扩展名: {file_ext}\n"
                f"大小: {file_size:,} 字节\n"
                f"状态: ✅ 加载成功"
            )
            
            # 注入渲染改进
            self.inject_rendering_improvements()
        else:
            self.info_label.setText("❌ 文件加载失败,请重试.")

    def export_pdf(self):
        """导出PDF文件"""
        if not self.page_loaded or not self.imported_file_path:
            self.info_label.setText("❌ 请先导入文件")
            return
        
        options = QFileDialog.Options()
        default_name = os.path.splitext(os.path.basename(self.imported_file_path))[0] + ".pdf"
        save_path, _ = QFileDialog.getSaveFileName(
            self, 
            "保存 PDF 文件", 
            os.path.join(self.last_directory, default_name), 
            "PDF Files (*.pdf);;All Files (*.*)", 
            options=options
        )

        if save_path:
            if not save_path.endswith('.pdf'):
                save_path += '.pdf'
            
            self.last_directory = os.path.dirname(save_path)
            
            # 显示导出进度
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self.info_label.setText("🔄 正在准备PDF导出...")
            self.export_button.setEnabled(False)
            
            # 延迟执行导出以确保所有渲染完成
            QTimer.singleShot(2000, lambda: self.perform_pdf_export(save_path))

    def preprocess_mht_file(self, mht_path):
        """预处理MHT文件以更好地保持样式和图片"""
        try:
            # 创建临时文件夹
            temp_dir = tempfile.mkdtemp()
            temp_html_path = os.path.join(temp_dir, "processed.html")
            
            # 尝试不同的编码来读取MHT文件
            content = None
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'utf-16', 'latin1']
            
            for encoding in encodings:
                try:
                    with open(mht_path, 'r', encoding=encoding, errors='ignore') as f:
                        content = f.read()
                    print(f"Successfully read MHT file with encoding: {encoding}")
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            
            if not content:
                print("Failed to read MHT file with any encoding, trying binary mode")
                # 如果所有编码都失败,尝试二进制模式
                with open(mht_path, 'rb') as f:
                    binary_content = f.read()
                # 尝试检测编码
                try:
                    content = binary_content.decode('utf-8', errors='replace')
                except:
                    content = binary_content.decode('latin1', errors='replace')
            
            # 解析MHT格式,提取HTML和图片
            html_content, images = self.extract_html_and_images_from_mht(content)
            
            if html_content:
                # 将图片保存到临时目录并更新HTML中的引用
                if images:
                    html_content = self.process_mht_images(html_content, images, temp_dir)
                
                # 确保HTML有正确的编码声明
                if '<meta charset=' not in html_content.lower() and '<meta http-equiv="content-type"' not in html_content.lower():
                    charset_meta = '<meta charset="UTF-8">\n'
                    if '<head>' in html_content:
                        html_content = html_content.replace('<head>', f'<head>\n{charset_meta}')
                    elif '<HEAD>' in html_content:
                        html_content = html_content.replace('<HEAD>', f'<HEAD>\n{charset_meta}')
                
                # 添加CSS来确保A4打印优化和高保真度转换
                enhanced_css = """
<style type="text/css">
/* A4打印优化的高保真度转换CSS */

/* 设置A4页面尺寸和边距 */
@page {
    size: A4 portrait;
    margin: 1cm 1.5cm;
}

/* 确保颜色和背景在PDF中正确显示 */
* {
    -webkit-print-color-adjust: exact !important;
    color-adjust: exact !important;
    print-color-adjust: exact !important;
    box-sizing: border-box !important;
}

/* 页面内容适配A4尺寸 */
body {
    margin: 0 !important;
    padding: 10px !important;
    font-family: "Microsoft YaHei", "SimSun", Arial, sans-serif !important;
    font-size: 12px !important;
    line-height: 1.3 !important;
    max-width: 100% !important;
    width: 100% !important;
}

/* 表格优化 - 适配A4宽度 */
table {
    width: 100% !important;
    border-collapse: collapse !important;
    margin: 0 auto 10px auto !important;
    page-break-inside: avoid !important;
    table-layout: auto !important;
}

/* 表格边框样式 */
table, td, th {
    border: 1px solid #000 !important;
}

td, th {
    padding: 4px 6px !important;
    vertical-align: top !important;
    word-wrap: break-word !important;
    word-break: break-all !important;
    font-size: 12px !important;
    line-height: 1.2 !important;
}

/* 表头样式 */
th {
    background-color: #f0f0f0 !important;
    font-weight: bold !important;
    text-align: center !important;
}

/* 图片优化 - 适配表格单元格 */
img {
    max-width: 120px !important;
    max-height: 150px !important;
    width: auto !important;
    height: auto !important;
    display: block !important;
    margin: 2px auto !important;
    page-break-inside: avoid !important;
    object-fit: contain !important;
}

/* 标题居中 */
h1, h2, h3 {
    text-align: center !important;
    margin: 10px 0 !important;
    font-size: 16px !important;
    font-weight: bold !important;
}

/* 文本对齐优化 */
.text-center, [align="center"] { 
    text-align: center !important; 
}
.text-left, [align="left"] { 
    text-align: left !important; 
}
.text-right, [align="right"] { 
    text-align: right !important; 
}

/* 特殊单元格样式保持 */
[bgcolor] { 
    background-color: attr(bgcolor) !important; 
}

/* 打印专用样式 */
@media print {
    /* 确保所有颜色在打印时保持 */
    * {
        -webkit-print-color-adjust: exact !important;
        color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    /* A4页面设置 */
    @page {
        size: A4 portrait;
        margin: 1cm 1.5cm;
    }
    
    /* 页面内容 */
    body {
        margin: 0 !important;
        padding: 5px !important;
        width: 100% !important;
        max-width: 100% !important;
    }
    
    /* 表格在打印时的优化 */
    table {
        width: 100% !important;
        page-break-inside: avoid !important;
        border-collapse: collapse !important;
    }
    
    tr {
        page-break-inside: avoid !important;
    }
    
    td, th {
        page-break-inside: avoid !important;
        border: 1px solid #000 !important;
        padding: 3px 5px !important;
        font-size: 11px !important;
    }
    
    /* 图片在打印时的优化 */
    img {
        max-width: 100px !important;
        max-height: 120px !important;
        page-break-inside: avoid !important;
    }
    
    /* 防止内容溢出 */
    * {
        overflow: visible !important;
    }
}

/* 响应式调整 - 确保内容适配页面 */
@media (max-width: 21cm) {
    body {
        font-size: 11px !important;
    }
    
    td, th {
        font-size: 11px !important;
        padding: 3px 4px !important;
    }
    
    img {
        max-width: 100px !important;
        max-height: 120px !important;
    }
}

</style>
"""
                
                # 在head标签中插入CSS
                if '<head>' in html_content:
                    html_content = html_content.replace('<head>', f'<head>\n{enhanced_css}')
                elif '<HEAD>' in html_content:
                    html_content = html_content.replace('<HEAD>', f'<HEAD>\n{enhanced_css}')
                else:
                    # 如果没有head标签,在html标签后添加
                    if '<html' in html_content:
                        insert_pos = html_content.find('>', html_content.find('<html')) + 1
                        html_content = html_content[:insert_pos] + f'\n<head>\n{enhanced_css}\n</head>\n' + html_content[insert_pos:]
                
                # 写入处理后的HTML文件
                with open(temp_html_path, 'w', encoding='utf-8', errors='replace') as f:
                    f.write(html_content)
                
                return temp_html_path
            
            return None
            
        except Exception as e:
            print(f"Error preprocessing MHT file: {e}")
            return None

    def extract_html_and_images_from_mht(self, content):
        """从MHT内容中提取HTML部分和图片"""
        try:
            lines = content.split('\n')
            html_content = None
            images = {}
            
            i = 0
            while i < len(lines):
                line = lines[i]
                
                # 查找HTML内容部分
                if 'Content-Type: text/html' in line or 'content-type: text/html' in line.lower():
                    html_content = self.extract_section_content(lines, i)
                    print("Found HTML section")
                
                # 查找图片内容部分
                elif ('Content-Type: image/' in line or 'content-type: image/' in line.lower()):
                    # 提取Content-Location
                    content_location = None
                    content_transfer_encoding = None
                    
                    j = i
                    while j < len(lines) and lines[j].strip() != '':
                        if 'Content-Location:' in lines[j]:
                            content_location = lines[j].split(':', 1)[1].strip()
                        elif 'Content-Transfer-Encoding:' in lines[j]:
                            content_transfer_encoding = lines[j].split(':', 1)[1].strip().lower()
                        j += 1
                    
                    if content_location:
                        # 提取图片数据
                        image_data = self.extract_section_content(lines, i, is_binary=True)
                        if image_data and content_transfer_encoding == 'base64':
                            try:
                                # 解码base64图片数据
                                decoded_image = base64.b64decode(image_data.replace('\n', '').replace('\r', ''))
                                images[content_location] = decoded_image
                                print(f"Found image: {content_location}")
                            except Exception as e:
                                print(f"Error decoding image {content_location}: {e}")
                
                i += 1
            
            # 如果通过HTML section方法找到了内容,进行解码
            if html_content:
                # 检查是否需要quoted-printable解码
                if '=E' in html_content and '=9' in html_content:  # quoted-printable的特征
                    try:
                        decoded_bytes = quopri.decodestring(html_content.encode('latin1'))
                        for encoding in ['utf-8', 'gbk', 'gb2312', 'gb18030']:
                            try:
                                html_content = decoded_bytes.decode(encoding)
                                print(f"Successfully decoded HTML with {encoding}")
                                break
                            except UnicodeDecodeError:
                                continue
                        else:
                            html_content = decoded_bytes.decode('utf-8', errors='replace')
                    except Exception as e:
                        print(f"Error decoding quoted-printable HTML: {e}")
            
            # 如果没有找到HTML section,尝试简单搜索
            if not html_content:
                html_start_patterns = ['<html', '<HTML', '<!DOCTYPE', '<!doctype']
                for pattern in html_start_patterns:
                    start_pos = content.find(pattern)
                    if start_pos != -1:
                        html_content = content[start_pos:]
                        # 查找可能的结束boundary
                        boundary_patterns = ['------=', '----boundary', '--======']
                        for boundary in boundary_patterns:
                            boundary_pos = html_content.find(boundary)
                            if boundary_pos != -1:
                                html_content = html_content[:boundary_pos]
                                break
                        print("Found HTML using simple search")
                        break
            
            return html_content, images
            
        except Exception as e:
            print(f"Error extracting HTML and images from MHT: {e}")
            return None, {}

    def extract_section_content(self, lines, start_index, is_binary=False):
        """提取MHT section的内容"""
        try:
            # 跳过头部信息到空行
            i = start_index + 1
            while i < len(lines) and lines[i].strip() != '':
                i += 1
            
            # 跳过空行
            i += 1
            
            # 收集内容直到下一个boundary
            content_lines = []
            while i < len(lines):
                line = lines[i]
                if (line.startswith('------=') or 
                    line.startswith('----boundary') or
                    line.startswith('--======')):
                    break
                content_lines.append(line)
                i += 1
            
            return '\n'.join(content_lines) if content_lines else None
            
        except Exception as e:
            print(f"Error extracting section content: {e}")
            return None

    def process_mht_images(self, html_content, images, temp_dir):
        """处理MHT中的图片,将其保存为本地文件并更新HTML引用"""
        try:
            # 为每个图片创建本地文件
            image_mapping = {}
            
            for location, image_data in images.items():
                # 提取文件名和扩展名
                filename = os.path.basename(location)
                if not filename or '.' not in filename:
                    # 根据图片数据推测格式
                    if image_data.startswith(b'\xff\xd8\xff'):
                        filename = f"image_{len(image_mapping)}.jpg"
                    elif image_data.startswith(b'\x89PNG'):
                        filename = f"image_{len(image_mapping)}.png"
                    elif image_data.startswith(b'GIF'):
                        filename = f"image_{len(image_mapping)}.gif"
                    else:
                        filename = f"image_{len(image_mapping)}.jpg"
                
                # 保存图片到临时目录
                image_path = os.path.join(temp_dir, filename)
                with open(image_path, 'wb') as f:
                    f.write(image_data)
                
                image_mapping[location] = image_path
                print(f"Saved image: {filename}")
            
            # 更新HTML中的图片引用
            for original_location, local_path in image_mapping.items():
                # 尝试多种可能的引用格式
                patterns_to_replace = [
                    f'src="{original_location}"',
                    f"src='{original_location}'",
                    f'src={original_location}',
                    original_location
                ]
                
                # 使用file://协议的本地路径
                local_url = QUrl.fromLocalFile(local_path).toString()
                
                for pattern in patterns_to_replace:
                    if pattern in html_content:
                        html_content = html_content.replace(pattern, f'src="{local_url}"')
                        print(f"Replaced image reference: {pattern}")
            
            return html_content
            
        except Exception as e:
            print(f"Error processing MHT images: {e}")
            return html_content

    def on_page_loaded(self, ok):
        """页面加载完成回调"""
        self.progress_bar.setVisible(False)
        if ok:
            self.page_loaded = True
            self.export_button.setEnabled(True)
            
            # 更新状态信息
            file_name = os.path.basename(self.imported_file_path)
            file_dir = os.path.dirname(self.imported_file_path)
            file_ext = os.path.splitext(self.imported_file_path)[1]
            file_size = os.path.getsize(self.imported_file_path)
            
            self.info_label.setText(
                f"文件名: {file_name}\n"
                f"位置: {file_dir}\n"
                f"扩展名: {file_ext}\n"
                f"大小: {file_size:,} 字节\n"
                f"状态: ✓ 加载成功 - 可以导出"
            )
            
            # 注入额外的CSS来进一步改善渲染
            self.inject_rendering_improvements()
        else:
            self.page_loaded = False
            self.export_button.setEnabled(False)
            self.info_label.setText("❌ Failed to load file. Please try again.")

    def inject_rendering_improvements(self):
        """注入A4打印优化的JavaScript"""
        js_code = """
// 等待页面完全加载
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyA4PrintOptimizations);
} else {
    applyA4PrintOptimizations();
}

function applyA4PrintOptimizations() {
    console.log('Applying A4 print optimizations...');
    
    // 保护原有字体样式
    function preserveOriginalFontStyles() {
        console.log('Preserving original font styles...');
        
        // 首先保护所有已有内联样式的元素
        var allElements = document.querySelectorAll('*');
        allElements.forEach(function(element) {
            var style = element.getAttribute('style');
            if (style) {
                // 检查是否包含字体相关的样式
                if (style.includes('font-size') || style.includes('font-family') || 
                    style.includes('font-weight') || style.includes('font-style') ||
                    style.includes('fontSize') || style.includes('fontFamily') ||
                    style.includes('fontWeight') || style.includes('fontStyle')) {
                    element.setAttribute('data-preserve-font', 'true');
                    console.log('Protected element with font style:', element.tagName, style);
                }
            }
            
            // 检查计算样式中的字体设置
            var computedStyle = window.getComputedStyle(element);
            var defaultFontSize = '16px'; // 浏览器默认字体大小
            
            // 如果元素的字体大小不是默认值,说明被特别设置过
            if (computedStyle.fontSize && computedStyle.fontSize !== defaultFontSize) {
                element.setAttribute('data-preserve-font', 'true');
                element.setAttribute('data-original-font-size', computedStyle.fontSize);
                console.log('Protected element with computed font size:', element.tagName, computedStyle.fontSize);
            }
            
            // 保护特殊的字体家族设置
            if (computedStyle.fontFamily && computedStyle.fontFamily !== 'Times') {
                element.setAttribute('data-preserve-font', 'true');
                element.setAttribute('data-original-font-family', computedStyle.fontFamily);
            }
            
            // 保护字体粗细设置
            if (computedStyle.fontWeight && computedStyle.fontWeight !== 'normal' && computedStyle.fontWeight !== '400') {
                element.setAttribute('data-preserve-font', 'true');
                element.setAttribute('data-original-font-weight', computedStyle.fontWeight);
            }
        });
        
        // 额外保护表格单元格的字体样式
        var tableCells = document.querySelectorAll('td, th');
        tableCells.forEach(function(cell) {
            cell.setAttribute('data-preserve-font', 'true');
            var computedStyle = window.getComputedStyle(cell);
            if (computedStyle.fontSize) {
                cell.setAttribute('data-original-font-size', computedStyle.fontSize);
            }
            if (computedStyle.fontFamily) {
                cell.setAttribute('data-original-font-family', computedStyle.fontFamily);
            }
            if (computedStyle.fontWeight) {
                cell.setAttribute('data-original-font-weight', computedStyle.fontWeight);
            }
            console.log('Protected table cell font:', cell.tagName, computedStyle.fontSize, computedStyle.fontFamily);
        });
    }
    
    // 清理空白表格行
    function removeEmptyTableRows() {
        console.log('Removing empty table rows...');
        var tables = document.querySelectorAll('table');
        tables.forEach(function(table) {
            var rows = table.querySelectorAll('tr');
            rows.forEach(function(row) {
                // 检查是否为空行
                var cells = row.querySelectorAll('td, th');
                var isEmpty = true;
                
                for (var i = 0; i < cells.length; i++) {
                    var cellText = cells[i].textContent.trim();
                    var cellHTML = cells[i].innerHTML.trim();
                    
                    // 如果有文字内容或有意义的HTML内容(不只是空格、换行符、&nbsp;)
                    if (cellText && cellText !== '' && cellText !== '\\u00A0') {
                        isEmpty = false;
                        break;
                    }
                    
                    // 检查是否有图片或其他有意义的元素
                    if (cells[i].querySelector('img, input, select, textarea')) {
                        isEmpty = false;
                        break;
                    }
                    
                    // 检查HTML内容(排除只有空白字符的情况)
                    var cleanHTML = cellHTML.replace(/&nbsp;/g, '').replace(/\\s/g, '');
                    if (cleanHTML && cleanHTML !== '') {
                        isEmpty = false;
                        break;
                    }
                }
                
                // 如果是空行,移除它
                if (isEmpty) {
                    console.log('Removing empty row');
                    row.remove();
                }
            });
        });
    }
    
    // 优化表格适配A4纸张
    function optimizeTablesForA4() {
        var tables = document.querySelectorAll('table');
        tables.forEach(function(table) {
            // 设置表格基本样式
            table.style.width = '100%';
            table.style.borderCollapse = 'collapse';
            table.style.margin = '0 auto 10px auto';
            table.style.tableLayout = 'auto';
            
            // 优化单元格
            var cells = table.querySelectorAll('td, th');
            cells.forEach(function(cell) {
                cell.style.padding = '4px 6px';
                cell.style.verticalAlign = 'top';
                cell.style.wordWrap = 'break-word';
                
                // 检查是否需要保护原有字体样式
                var preserveFont = cell.getAttribute('data-preserve-font') === 'true';
                
                if (preserveFont) {
                    // 恢复保存的原始字体设置
                    var originalFontSize = cell.getAttribute('data-original-font-size');
                    var originalFontFamily = cell.getAttribute('data-original-font-family');
                    var originalFontWeight = cell.getAttribute('data-original-font-weight');
                    
                    if (originalFontSize) {
                        cell.style.fontSize = originalFontSize;
                        console.log('Restored font size:', originalFontSize, 'for', cell.tagName);
                    }
                    
                    if (originalFontFamily) {
                        cell.style.fontFamily = originalFontFamily;
                    }
                    
                    if (originalFontWeight) {
                        cell.style.fontWeight = originalFontWeight;
                    }
                } else {
                    // 只在没有保护标记且没有现有字体大小时才设置默认值
                    if (!cell.style.fontSize && !cell.getAttribute('style')?.includes('font-size')) {
                        cell.style.fontSize = '12px';
                    }
                    
                    // 只在没有保护标记且没有现有行高时才设置默认值
                    if (!cell.style.lineHeight && !cell.getAttribute('style')?.includes('line-height')) {
                        cell.style.lineHeight = '1.2';
                    }
                }
                
                cell.style.border = '1px solid #000';
            });
            
            // 特殊处理表头
            var headers = table.querySelectorAll('th');
            headers.forEach(function(th) {
                th.style.backgroundColor = '#f0f0f0';
                
                // 只在没有保护标记时才设置默认字体粗细和对齐
                var preserveFont = th.getAttribute('data-preserve-font') === 'true';
                if (!preserveFont) {
                    th.style.fontWeight = 'bold';
                    th.style.textAlign = 'center';
                }
            });
        });
    }
    
    // 优化图片适配表格和A4纸张
    function optimizeImagesForA4() {
        var images = document.querySelectorAll('img');
        images.forEach(function(img) {
            // 检查图片是否在表格中
            var isInTable = img.closest('table') !== null;
            
            if (isInTable) {
                // 表格中的图片使用较小尺寸
                img.style.maxWidth = '120px';
                img.style.maxHeight = '150px';
            } else {
                // 表格外的图片可以稍大一些
                img.style.maxWidth = '200px';
                img.style.maxHeight = '250px';
            }
            
            img.style.width = 'auto';
            img.style.height = 'auto';
            img.style.display = 'block';
            img.style.margin = '2px auto';
            img.style.objectFit = 'contain';
            
            // 图片加载错误处理
            img.onerror = function() {
                this.style.border = '1px dashed #ccc';
                this.style.background = '#f9f9f9';
                this.style.minWidth = '50px';
                this.style.minHeight = '50px';
                this.alt = '图片加载失败';
            };
        });
    }
    
    // 优化页面布局适配A4
    function optimizePageLayoutForA4() {
        var body = document.body;
        if (body) {
            body.style.margin = '0';
            body.style.padding = '10px';
            body.style.maxWidth = '100%';
            body.style.width = '100%';
            
            // 检查是否需要保护原有字体样式
            var preserveBodyFont = body.getAttribute('data-preserve-font') === 'true';
            
            // 只在没有保护标记且没有现有字体设置时才应用默认字体
            if (!preserveBodyFont && !body.style.fontFamily && !body.getAttribute('style')?.includes('font-family')) {
                body.style.fontFamily = '"Microsoft YaHei", "SimSun", Arial, sans-serif';
            }
            
            // 只在没有保护标记且没有现有字体大小时才应用默认大小
            if (!preserveBodyFont && !body.style.fontSize && !body.getAttribute('style')?.includes('font-size')) {
                body.style.fontSize = '12px';
            }
            
            // 只在没有保护标记且没有现有行高时才应用默认行高
            if (!preserveBodyFont && !body.style.lineHeight && !body.getAttribute('style')?.includes('line-height')) {
                body.style.lineHeight = '1.3';
            }
        }
        
        // 优化标题 - 保留原有样式,只补充必要的居中和间距
        var headings = document.querySelectorAll('h1, h2, h3');
        headings.forEach(function(h) {
            h.style.textAlign = 'center';
            h.style.margin = '10px 0';
            
            // 检查是否需要保护原有字体样式
            var preserveHeadingFont = h.getAttribute('data-preserve-font') === 'true';
            
            // 只在没有保护标记且没有现有字体大小时才设置默认大小
            if (!preserveHeadingFont && !h.style.fontSize && !h.getAttribute('style')?.includes('font-size')) {
                h.style.fontSize = '16px';
            }
            
            // 只在没有保护标记且没有现有字体粗细时才设置粗体
            if (!preserveHeadingFont && !h.style.fontWeight && !h.getAttribute('style')?.includes('font-weight')) {
                h.style.fontWeight = 'bold';
            }
        });
    }
    
    // 确保打印颜色保真度
    function ensurePrintColorFidelity() {
        var style = document.createElement('style');
        style.type = 'text/css';
        style.innerHTML = `
            /* A4打印专用样式 */
            @media print {
                @page {
                    size: A4 portrait;
                    margin: 1cm 1.5cm;
                }
                
                * {
                    -webkit-print-color-adjust: exact !important;
                    color-adjust: exact !important;
                }
                
                body {
                    margin: 0 !important;
                    padding: 5px !important;
                    width: 100% !important;
                }
                
                table {
                    width: 100% !important;
                    page-break-inside: avoid !important;
                }
                
                tr {
                    page-break-inside: avoid !important;
                }
                
                td, th {
                    page-break-inside: avoid !important;
                    padding: 3px 5px !important;
                }
                
                img {
                    max-width: 100px !important;
                    max-height: 120px !important;
                    page-break-inside: avoid !important;
                }
            }
        `;
        
        if (document.head) {
            document.head.appendChild(style);
        }
    }
    
    // 执行所有A4优化
    preserveOriginalFontStyles();  // 首先保护原始字体样式
    removeEmptyTableRows();  // 清理空白表格行
    optimizeTablesForA4();
    optimizeImagesForA4();
    optimizePageLayoutForA4();
    ensurePrintColorFidelity();
    
    console.log('A4 print optimizations applied successfully');
    
    // 计算和显示页面信息
    setTimeout(function() {
        var pageHeight = document.body.scrollHeight;
        var a4Height = 297 * 3.78; // A4高度转换为像素(约1122px)
        console.log('Page height: ' + pageHeight + 'px, A4 height: ~' + a4Height + 'px');
        
        if (pageHeight > a4Height * 0.9) {
            console.log('Warning: Content may exceed A4 page size');
        }
    }, 500);
}
"""
        
        self.web_view.page().runJavaScript(js_code)

    def export_pdf(self):
        """导出PDF文件"""
        if not hasattr(self, 'imported_file_path') or not self.imported_file_path:
            self.info_label.setText("❌ Please import a file first.")
            return
            
        if not self.page_loaded:
            self.info_label.setText("⏳ Please wait for the page to load completely before exporting.")
            return
                
        options = QFileDialog.Options()
        default_name = os.path.splitext(os.path.basename(self.imported_file_path))[0] + ".pdf"
        save_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Save PDF File", 
            os.path.join(self.last_directory, default_name), 
            "PDF Files (*.pdf);;All Files (*.*)", 
            options=options
        )

        if save_path:
            if not save_path.endswith('.pdf'):
                save_path += '.pdf'
            
            self.last_directory = os.path.dirname(save_path)
            
            # 显示导出进度
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self.info_label.setText("🔄 Preparing PDF export...")
            self.export_button.setEnabled(False)
            
            # 延迟执行导出以确保所有渲染完成
            QTimer.singleShot(2000, lambda: self.perform_pdf_export(save_path))

    def perform_pdf_export(self, save_path):
        """执行A4优化的PDF导出"""
        try:
            # A4打印优化的最终样式调整
            final_js = """
            // A4打印优化最终调整
            console.log('Applying final A4 print optimizations...');
            
            // 动态应用样式,避免覆盖保护的字体设置
            function applyFinalStyles() {
                // 应用基本的A4页面样式
                var finalStyle = document.createElement('style');
                finalStyle.innerHTML = `
                    /* A4打印专用最终样式 */
                    @page {
                        size: A4 portrait;
                        margin: 1cm 1.5cm;
                    }
                    
                    /* 确保内容适配A4页面 */
                    body {
                        margin: 0 !important;
                        padding: 10px !important;
                        background: white !important;
                        max-width: 100% !important;
                    }
                    
                    /* 表格A4适配 */
                    table {
                        width: 100% !important;
                        border-collapse: collapse !important;
                        margin: 0 auto 8px auto !important;
                        page-break-inside: avoid !important;
                        table-layout: auto !important;
                    }
                    
                    /* 图片A4适配 */
                    img {
                        max-width: 120px !important;
                        max-height: 150px !important;
                        width: auto !important;
                        height: auto !important;
                        display: block !important;
                        margin: 2px auto !important;
                        page-break-inside: avoid !important;
                        object-fit: contain !important;
                    }
                `;
                
                if (document.head) {
                    document.head.appendChild(finalStyle);
                }
                
                // 为没有保护标记的单元格应用基本样式
                var cells = document.querySelectorAll('td, th');
                cells.forEach(function(cell) {
                    // 始终应用边框和布局样式
                    cell.style.border = '1px solid #000';
                    cell.style.padding = '4px 6px';
                    cell.style.wordWrap = 'break-word';
                    cell.style.verticalAlign = 'top';
                });
                
                // 为没有保护标记的表头应用样式
                var headers = document.querySelectorAll('th');
                headers.forEach(function(th) {
                    th.style.backgroundColor = '#f0f0f0';
                    
                    var preserveFont = th.getAttribute('data-preserve-font') === 'true';
                    if (!preserveFont) {
                        th.style.fontWeight = 'bold';
                        th.style.textAlign = 'center';
                    }
                });
            }
            
            applyFinalStyles();
            """
            
            self.web_view.page().runJavaScript(final_js)
            
            # 等待JavaScript执行和布局计算完成后导出
            QTimer.singleShot(3000, lambda: self.do_pdf_export(save_path))
            
        except Exception as e:
            self.handle_export_error(f"Export preparation failed: {e}")

    def do_pdf_export(self, save_path):
        """实际执行PDF导出 - A4优化版本"""
        try:
            self.info_label.setText("📄 Generating PDF with A4 optimization...")
            
            # 使用A4优化设置进行PDF导出
            self.web_view.page().printToPdf(save_path)
            
            # 等待导出完成
            QTimer.singleShot(4000, lambda: self.on_export_complete(save_path))
            
        except Exception as e:
            self.handle_export_error(f"PDF export failed: {e}")

    def on_export_complete(self, save_path):
        """导出完成处理"""
        try:
            if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                self.progress_bar.setVisible(False)
                self.export_button.setEnabled(True)
                
                file_size = os.path.getsize(save_path)
                self.info_label.setText(f"✅ PDF exported successfully!\nLocation: {save_path}\nSize: {file_size:,} bytes")
                
                # 打开文件所在文件夹并选中文件
                subprocess.Popen(f'explorer /select,"{os.path.abspath(save_path)}"', shell=True)
            else:
                self.handle_export_error("PDF file was not created or is empty")
                
        except Exception as e:
            self.handle_export_error(f"Post-export processing failed: {e}")

    def handle_export_error(self, error_msg):
        """处理导出错误"""
        self.progress_bar.setVisible(False)
        self.export_button.setEnabled(True)
        self.info_label.setText(f"❌ Export failed: {error_msg}")
        print(f"Export error: {error_msg}")

app = QApplication(sys.argv)
window = HTMLtoPDFConverter()
window.show()
app.exec_()
