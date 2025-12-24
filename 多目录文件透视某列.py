import sys
import os
import pandas as pd
from pathlib import Path
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *


class FileDropWidget(QListWidget):
    directories_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setStyleSheet("""
            QListWidget {
                border: 2px dashed #aaa;
                border-radius: 5px;
                background-color: #f8f9fa;
                min-height: 80px;
                max-height: 120px;
            }
            QListWidget::item {
                padding: 5px;
            }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        directories = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isdir(path):
                directories.append(path)

        if directories:
            self.directories_dropped.emit(directories)
        event.accept()


import re

class AnalysisWorker(QThread):
    progress_updated = pyqtSignal(str)
    file_result_ready = pyqtSignal(dict)  # 修改为发送单文件结果
    finished_analysis = pyqtSignal(dict)  # 传递汇总数据
    error_occurred = pyqtSignal(str)

    def __init__(self, directories, amount_column, custom_headers):
        super().__init__()
        self.directories = directories
        self.amount_column = amount_column
        self.custom_headers = custom_headers
        self.is_running = True

    def _extract_monthly_card(self, filename):
        """从文件名提取月结卡号"""
        # 匹配10位数字
        pattern = r'\d{10}'
        match = re.search(pattern, filename)
        if match:
            return match.group(0)
        return ""

    def _preprocess_df(self, df):
        """预处理DataFrame，移除合计列和行"""
        # 移除名为"合 计"或"合计"的列
        cols_to_drop = [c for c in df.columns if str(c).strip() in ['合 计', '合计']]
        if cols_to_drop:
            df = df.drop(columns=cols_to_drop)
            
        # 移除包含"合 计"或"合计"的行
        # 只需要检查Object类型的列
        for col in df.columns:
            if df[col].dtype == 'object':
                # 检查是否包含特定关键词
                mask = df[col].astype(str).str.contains('合 计|合计', regex=True, na=False)
                df = df[~mask]
                
        return df

    def run(self):
        try:
            total_amount_all = 0
            total_files_all = 0
            
            for directory in self.directories:
                if not self.is_running:
                    break

                self.progress_updated.emit(f"正在分析目录: {directory}")
                
                # 读取所有Excel文件
                for excel_file in Path(directory).glob("*.xlsx"):
                    if not self.is_running:
                        break
                    try:
                        df = pd.read_excel(excel_file, header=0)
                        if self.custom_headers is not None and len(self.custom_headers) == len(df.columns):
                            df.columns = self.custom_headers

                        # 预处理数据，移除合计行/列
                        df = self._preprocess_df(df)

                        if self.amount_column in df.columns:
                            # 转换金额列为数值类型
                            df[self.amount_column] = pd.to_numeric(df[self.amount_column], errors='coerce')
                            file_total = df[self.amount_column].sum()
                            
                            # 发送单文件结果
                            result = {
                                'directory': str(directory),
                                'filename': excel_file.name,
                                'monthly_card': self._extract_monthly_card(excel_file.name),
                                'amount': float(file_total)
                            }
                            self.file_result_ready.emit(result)
                            
                            total_amount_all += file_total
                            total_files_all += 1
                    except Exception as e:
                        print(f"读取文件 {excel_file} 失败: {e}")

                # 读取所有CSV文件
                for csv_file in Path(directory).glob("*.csv"):
                    if not self.is_running:
                        break
                    try:
                        # 尝试不同的编码读取 CSV
                        try:
                            df = pd.read_csv(csv_file, header=0, encoding='utf-8')
                        except UnicodeDecodeError:
                            df = pd.read_csv(csv_file, header=0, encoding='gbk')

                        if self.custom_headers is not None and len(self.custom_headers) == len(df.columns):
                            df.columns = self.custom_headers

                        # 预处理数据，移除合计行/列
                        df = self._preprocess_df(df)

                        if self.amount_column in df.columns:
                            df[self.amount_column] = pd.to_numeric(df[self.amount_column], errors='coerce')
                            file_total = df[self.amount_column].sum()
                            
                            # 发送单文件结果
                            result = {
                                'directory': str(directory),
                                'filename': csv_file.name,
                                'monthly_card': self._extract_monthly_card(csv_file.name),
                                'amount': float(file_total)
                            }
                            self.file_result_ready.emit(result)
                            
                            total_amount_all += file_total
                            total_files_all += 1
                    except Exception as e:
                        print(f"读取文件 {csv_file} 失败: {e}")

            summary = {
                'total_files': int(total_files_all),
                'total_amount': float(total_amount_all)
            }
            self.finished_analysis.emit(summary)

        except Exception as e:
            self.error_occurred.emit(str(e))

    def stop(self):
        self.is_running = False


class HeaderEditorDialog(QDialog):
    def __init__(self, headers, parent=None):
        super().__init__(parent)
        self.headers = headers
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("编辑表头")
        self.setMinimumWidth(500)

        layout = QVBoxLayout()

        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(1)
        self.table.setRowCount(len(self.headers))
        self.table.setHorizontalHeaderLabels(["表头名称"])
        self.table.verticalHeader().setVisible(False)

        for i, header in enumerate(self.headers):
            item = QTableWidgetItem(header)
            self.table.setItem(i, 0, item)

        layout.addWidget(self.table)

        # 按钮区域
        button_layout = QHBoxLayout()

        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(self.accept)

        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(btn_ok)
        button_layout.addWidget(btn_cancel)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def get_edited_headers(self):
        headers = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                headers.append(item.text())
        return headers


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.directories = []
        self.dataframes = []
        self.all_results = []  # 存储所有文件结果
        self.custom_headers = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("目录数据分析工具")
        self.setGeometry(100, 100, 1200, 800)

        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)

        # 标题
        title_label = QLabel("目录数据分析工具")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 拖放区域
        self.drop_widget = FileDropWidget(self)
        self.drop_widget.directories_dropped.connect(self.add_directories)
        
        # 顶部工具栏布局（提示文本 + 清空按钮）
        top_bar_layout = QHBoxLayout()
        
        self.drop_label = QLabel("拖拽目录到这里（支持Excel/CSV文件）")
        self.drop_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 14px;
            }
        """)
        top_bar_layout.addWidget(self.drop_label)
        top_bar_layout.addStretch()
        
        # 清空目录按钮
        self.clear_dirs_btn = QPushButton("清空目录")
        self.clear_dirs_btn.clicked.connect(self.clear_directories)
        self.clear_dirs_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                padding: 4px 8px;
                border-radius: 4px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        top_bar_layout.addWidget(self.clear_dirs_btn)

        drop_layout = QVBoxLayout()
        drop_layout.addLayout(top_bar_layout)
        drop_layout.addWidget(self.drop_widget)

        drop_group = QGroupBox("目录列表")
        drop_group.setLayout(drop_layout)
        # 设置固定高度，让它变小一点
        drop_group.setFixedHeight(180)
        main_layout.addWidget(drop_group)

        # 配置区域
        config_layout = QHBoxLayout()

        # 金额列输入
        amount_label = QLabel("金额列名：")
        self.amount_combo = QComboBox()
        self.amount_combo.setMinimumWidth(150)

        config_layout.addWidget(amount_label)
        config_layout.addWidget(self.amount_combo)

        # 表头编辑按钮
        self.header_edit_btn = QPushButton("编辑表头")
        self.header_edit_btn.clicked.connect(self.edit_headers)
        self.header_edit_btn.setEnabled(False)
        config_layout.addWidget(self.header_edit_btn)

        # 分析按钮
        self.analyze_btn = QPushButton("开始分析")
        self.analyze_btn.clicked.connect(self.analyze_data)
        self.analyze_btn.setEnabled(False)
        self.analyze_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        config_layout.addWidget(self.analyze_btn)

        # 导出按钮
        self.export_btn = QPushButton("导出结果")
        self.export_btn.clicked.connect(self.export_data)
        self.export_btn.setEnabled(False)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        config_layout.addWidget(self.export_btn)

        config_layout.addStretch()

        config_group = QGroupBox("分析配置")
        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)

        # 结果显示区域
        result_layout = QVBoxLayout()

        # 统计结果表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["目录", "文件名", "月结卡号", "金额"])
        self.result_table.horizontalHeader().setStretchLastSection(True)

        result_layout.addWidget(QLabel("统计结果："))
        result_layout.addWidget(self.result_table)

        # 总计行
        self.total_label = QLabel("")
        self.total_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #27ae60;
                padding: 10px;
                background-color: #ecf0f1;
                border-radius: 4px;
            }
        """)
        result_layout.addWidget(self.total_label)

        result_group = QGroupBox("分析结果")
        result_group.setLayout(result_layout)
        main_layout.addWidget(result_group, 1)  # 添加 stretch 因子 1，让它占据剩余空间

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("准备就绪")
        # 增大状态栏字体
        self.status_bar.setStyleSheet("QStatusBar { font-size: 14px; padding: 5px; }")

    def add_directories(self, directories):
        """添加目录到列表"""
        for directory in directories:
            if directory not in self.directories:
                self.directories.append(directory)

                # 获取目录中的Excel/CSV文件
                excel_files = list(Path(directory).glob("*.xlsx"))
                csv_files = list(Path(directory).glob("*.csv"))
                file_count = len(excel_files) + len(csv_files)

                item_text = f"{directory} ({file_count}个文件)"
                self.drop_widget.addItem(item_text)

        self.update_controls()
        self.status_bar.showMessage(f"已添加 {len(directories)} 个目录")

    def clear_directories(self):
        """清空目录列表"""
        self.directories.clear()
        self.drop_widget.clear()
        self.update_controls()
        self.drop_label.setText("拖拽目录到这里（支持Excel/CSV文件）")
        self.status_bar.showMessage("已清空目录列表")

    def update_controls(self):
        """更新控件状态"""
        has_directories = len(self.directories) > 0
        self.analyze_btn.setEnabled(has_directories)

        if has_directories:
            self.drop_label.setText(f"已添加 {len(self.directories)} 个目录，可继续拖拽添加")
            self.scan_first_file_headers()

    def scan_first_file_headers(self):
        """扫描第一个文件的表头"""
        for directory in self.directories:
            # 查找第一个Excel或CSV文件
            for ext in ['*.xlsx', '*.csv']:
                files = list(Path(directory).glob(ext))
                if files:
                    try:
                        file_path = files[0]
                        if file_path.suffix == '.csv':
                            df = pd.read_csv(file_path, nrows=1)
                        else:
                            df = pd.read_excel(file_path, nrows=1)

                        self.amount_combo.clear()
                        for header in df.columns:
                            header_str = str(header).strip()
                            if header_str not in ['合 计', '合计']:
                                self.amount_combo.addItem(header)

                        self.header_edit_btn.setEnabled(True)
                        return
                    except Exception as e:
                        print(f"读取文件头失败: {e}")
                        continue

    def edit_headers(self):
        """编辑表头"""
        current_headers = []
        for i in range(self.amount_combo.count()):
            current_headers.append(self.amount_combo.itemText(i))

        dialog = HeaderEditorDialog(current_headers, self)
        if dialog.exec_():
            edited_headers = dialog.get_edited_headers()
            self.custom_headers = edited_headers

            self.amount_combo.clear()
            for header in edited_headers:
                self.amount_combo.addItem(header)

    def analyze_data(self):
        """分析数据"""
        if not self.directories:
            QMessageBox.warning(self, "警告", "请先添加目录")
            return

        if self.amount_combo.currentText() == "":
            QMessageBox.warning(self, "警告", "请选择金额列")
            return

        amount_column = self.amount_combo.currentText()
        self.result_table.setRowCount(0)
        self.dataframes = []  # 清空旧数据
        self.all_results = []  # 清空旧结果
        
        # 禁用按钮防止重复点击
        self.analyze_btn.setEnabled(False)
        self.header_edit_btn.setEnabled(False)
        self.drop_widget.setEnabled(False)
        self.export_btn.setEnabled(False)
        
        # 创建并启动工作线程
        self.worker = AnalysisWorker(self.directories, amount_column, self.custom_headers)
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.file_result_ready.connect(self.update_result_table)
        self.worker.finished_analysis.connect(self.analysis_finished)
        self.worker.error_occurred.connect(self.analysis_error)
        
        self.worker.start()

    def update_progress(self, message):
        self.status_bar.showMessage(message)

    def update_result_table(self, result):
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        
        self.result_table.setItem(row, 0, QTableWidgetItem(result['directory']))
        self.result_table.setItem(row, 1, QTableWidgetItem(result['filename']))
        self.result_table.setItem(row, 2, QTableWidgetItem(result['monthly_card']))
        self.result_table.setItem(row, 3, QTableWidgetItem(f"{result['amount']:,.2f}"))
        
        self.all_results.append(result)

    def analysis_finished(self, summary):
        self.total_label.setText(
            f"总计 - 目录数: {len(self.directories)}, 文件数: {summary['total_files']}, 总金额: {summary['total_amount']:,.2f}")
        self.status_bar.showMessage(f"分析完成，共处理 {summary['total_files']} 个文件")
        
        # 恢复界面状态
        self.analyze_btn.setEnabled(True)
        self.header_edit_btn.setEnabled(True)
        self.drop_widget.setEnabled(True)
        self.export_btn.setEnabled(True)

    def analysis_error(self, error_msg):
        QMessageBox.critical(self, "错误", f"分析数据时出错: {error_msg}")
        self.status_bar.showMessage("分析出错")
        
        # 恢复界面状态
        self.analyze_btn.setEnabled(True)
        self.header_edit_btn.setEnabled(True)
        self.drop_widget.setEnabled(True)
        self.export_btn.setEnabled(True)

    def export_data(self):
        """导出数据"""
        if not self.all_results:
            QMessageBox.warning(self, "警告", "没有数据可导出")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(self, "导出结果", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                df = pd.DataFrame(self.all_results)
                # 重命名列
                df = df.rename(columns={
                    'directory': '目录',
                    'filename': '文件名',
                    'monthly_card': '月结卡号',
                    'amount': '金额'
                })
                # 调整列顺序
                df = df[['目录', '文件名', '月结卡号', '金额']]
                
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "成功", f"数据已导出到 {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {e}")

    def clear_data(self):
        """清空数据"""
        self.directories.clear()
        self.dataframes.clear()
        self.all_results.clear()
        self.drop_widget.clear()
        self.result_table.setRowCount(0)
        self.total_label.setText("")
        self.amount_combo.clear()
        self.header_edit_btn.setEnabled(False)
        self.analyze_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.drop_label.setText("拖拽目录到这里（支持Excel/CSV文件）")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # 设置应用样式
    app.setStyleSheet("""
        QMainWindow {
            background-color: #ecf0f1;
        }
        QGroupBox {
            font-weight: bold;
            border: 2px solid #bdc3c7;
            border-radius: 5px;
            margin-top: 10px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QTableWidget {
            border: 1px solid #bdc3c7;
            background-color: white;
        }
        QHeaderView::section {
            background-color: #34495e;
            color: white;
            padding: 4px;
            border: 1px solid #2c3e50;
        }
        QComboBox {
            padding: 5px;
            border: 1px solid #bdc3c7;
            border-radius: 3px;
        }
        QPushButton {
            padding: 8px 16px;
            border-radius: 4px;
            border: none;
            font-weight: bold;
        }
    """)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())