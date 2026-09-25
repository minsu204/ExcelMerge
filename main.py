import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QFileDialog, QLabel,
    QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDropEvent
from merger import ExcelMerger


class ExcelMergeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.init_ui()
        self.setAcceptDrops(True)

    def init_ui(self):
        """GUI 초기화"""
        self.setWindowTitle("ExcelMerge - 엑셀 파일 병합 도구")
        self.setGeometry(100, 100, 600, 500)

        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 메인 레이아웃
        main_layout = QVBoxLayout()

        # 제목
        title_label = QLabel("엑셀 파일 병합 도구")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # 안내 메시지
        info_label = QLabel("💡 팁: 파일을 창에 드래그하여 업로드할 수 있습니다!")
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        main_layout.addWidget(info_label)

        # 파일 목록 레이블
        file_list_label = QLabel("선택된 파일:")
        main_layout.addWidget(file_list_label)

        # 파일 목록 (ListWidget)
        self.file_list = QListWidget()
        main_layout.addWidget(self.file_list)

        # 버튼 레이아웃
        button_layout = QHBoxLayout()

        # 파일 추가 버튼
        add_btn = QPushButton("파일 추가 (+)")
        add_btn.clicked.connect(self.add_files)
        button_layout.addWidget(add_btn)

        # 파일 제거 버튼
        remove_btn = QPushButton("제거 (-)")
        remove_btn.clicked.connect(self.remove_selected_file)
        button_layout.addWidget(remove_btn)

        # 모두 제거 버튼
        clear_btn = QPushButton("모두 제거")
        clear_btn.clicked.connect(self.clear_all_files)
        button_layout.addWidget(clear_btn)

        main_layout.addLayout(button_layout)

        # Merge 버튼 (강조)
        merge_btn = QPushButton("Merge 시작")
        merge_btn.setStyleSheet(
            "background-color: #4CAF50; color: white; "
            "font-weight: bold; padding: 10px; font-size: 14px;"
        )
        merge_btn.clicked.connect(self.merge_files)
        main_layout.addWidget(merge_btn)

        # 상태 레이블
        self.status_label = QLabel("상태: 준비 완료")
        main_layout.addWidget(self.status_label)

        central_widget.setLayout(main_layout)

    def add_files(self):
        """파일 추가 다이얼로그"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        for file_path in file_paths:
            if file_path not in self.selected_files:
                self.selected_files.append(file_path)
                # 파일명만 표시하되, 전체 경로는 data로 저장
                item = QListWidgetItem(os.path.basename(file_path))
                item.setData(Qt.UserRole, file_path)
                self.file_list.addItem(item)

        if file_paths:
            self.status_label.setText(f"상태: {len(self.selected_files)}개 파일 선택됨")

    def remove_selected_file(self):
        """선택된 파일 제거"""
        current_row = self.file_list.currentRow()
        if current_row >= 0:
            self.file_list.takeItem(current_row)
            self.selected_files.pop(current_row)
            self.status_label.setText(f"상태: {len(self.selected_files)}개 파일 선택됨")

    def clear_all_files(self):
        """모든 파일 제거"""
        self.file_list.clear()
        self.selected_files.clear()
        self.status_label.setText("상태: 준비 완료")

    def merge_files(self):
        """파일 병합 시작"""
        if not self.selected_files:
            QMessageBox.warning(self, "경고", "병합할 파일을 선택하세요!")
            return

        if len(self.selected_files) < 2:
            QMessageBox.warning(self, "경고", "2개 이상의 파일을 업로드해주세요!")
            return

        self.status_label.setText("상태: 파일 검증 중...")

        try:
            # 파일 병합
            merger = ExcelMerger(self.selected_files)
            merger.merge()

            # 저장 위치 선택
            file_path, selected_filter = QFileDialog.getSaveFileName(
                self,
                "병합된 파일 저장",
                "merged_result.xlsx",
                "Excel 2007+ (*.xlsx);;Excel 97-2003 (*.xls);;All Files (*)"
            )

            if file_path:
                # 확장자가 없으면 필터에 따라 추가
                if not file_path.endswith(('.xls', '.xlsx')):
                    if '*.xls' in selected_filter:
                        file_path += '.xls'
                    else:
                        file_path += '.xlsx'
                
                merger.save(file_path)
                self.status_label.setText("상태: 병합 완료!")
                QMessageBox.information(self, "성공", f"파일이 저장되었습니다!\n{file_path}")
            else:
                self.status_label.setText("상태: 취소됨")

        except ValueError as e:
            QMessageBox.critical(self, "오류", str(e))
            self.status_label.setText("상태: 오류 발생")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"병합 중 오류 발생:\n{str(e)}")
            self.status_label.setText("상태: 오류 발생")

    def dragEnterEvent(self, event):
        """드래그 진입 시"""
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        """파일 드롭 시"""
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        
        excel_files = []
        for file in files:
            if file.lower().endswith(('.xlsx', '.xls')):
                excel_files.append(file)
        
        if not excel_files:
            QMessageBox.warning(self, "경고", "엑셀 파일(.xlsx, .xls)만 지원합니다!")
            return
        
        # 파일 추가
        for file_path in excel_files:
            if file_path not in self.selected_files:
                self.selected_files.append(file_path)
                # 파일명만 표시하되, 전체 경로는 data로 저장
                item = QListWidgetItem(os.path.basename(file_path))
                item.setData(Qt.UserRole, file_path)
                self.file_list.addItem(item)
        
        self.status_label.setText(f"상태: {len(self.selected_files)}개 파일 선택됨")
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = ExcelMergeApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
