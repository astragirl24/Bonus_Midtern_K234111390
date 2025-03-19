import sys
import subprocess
from PyQt6 import QtCore, QtGui, QtWidgets
import pandas as pd
import plotly.express as px
import webbrowser

from bonus_midterm_K234111390.MainWindow import Ui_MainWindow


class MainWindowEx(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Khởi tạo giao diện từ Ui_MainWindow

        # Kết nối sự kiện cho các nút
        self.pushButton_2.clicked.connect(self.select_file)
        self.pushButton.clicked.connect(self.create_sunburst_chart)
        self.pushButtonSaveChart.clicked.connect(self.savechart)

    def select_file(self):
        # Mở hộp thoại chọn file
        file, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Chọn file Excel", "", "Excel Files (*.xlsx)")
        if file:
            self.lineEdit.setText(file)  # Hiển thị đường dẫn file vào lineEdit

    def create_sunburst_chart(self):
        file_path = self.lineEdit.text()  # Lấy đường dẫn file từ lineEdit
        if file_path:
            # Đọc dữ liệu từ file Excel
            df = pd.read_excel(file_path)

            # Lọc bỏ các dòng có giá trị thiếu trong các cột cần thiết
            df = df.dropna(subset=['Học Kỳ', 'Mã học phần', 'Tên học phần'])

            # Chuyển đổi cột 'Tín Chỉ' thành số, thay thế các giá trị không hợp lệ bằng NaN
            df['Tín Chỉ'] = pd.to_numeric(df['Tín Chỉ'], errors='coerce')

            # Loại bỏ các dòng có giá trị NaN trong cột 'Tín Chỉ'
            df = df.dropna(subset=['Tín Chỉ'])

            # Chuyển đổi giá trị của cột "Học Kỳ" thành "Học Kỳ 1", "Học Kỳ 2", ...
            df['Học Kỳ'] = df['Học Kỳ'].apply(lambda x: f'Học Kỳ {int(x)}')

            # Tạo biểu đồ Sunburst
            self.fig = px.sunburst(df, path=['Học Kỳ', 'Bắt buộc/tự chọn', 'Tên học phần'], values='Tín Chỉ',
                                   title='Biểu đồ Sunburst theo Học Kỳ, Bắt buộc/Tự chọn và Tên học phần')

            # Lưu kết quả ra file HTML
            self.fig.write_html('sunburst_chart.html')
            print("Biểu đồ Sunburst đã được tạo và lưu thành công!")

            # Hỏi người dùng có muốn hiển thị biểu đồ không
            reply = QtWidgets.QMessageBox.question(self, 'Xác nhận',
                                                   'Bạn có muốn hiển thị biểu đồ Sunburst không?',
                                                   QtWidgets.QMessageBox.StandardButton.Yes |
                                                   QtWidgets.QMessageBox.StandardButton.No,
                                                   QtWidgets.QMessageBox.StandardButton.Yes)
            if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                # Mở file HTML trong trình duyệt Chrome
                self.open_in_chrome('sunburst_chart.html')

    def open_in_chrome(self, file_path):
        # Đường dẫn đến trình duyệt Chrome (Cập nhật đường dẫn nếu cần)
        chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"  # Dành cho macOS
        # chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # Dành cho Windows

        # Mở file HTML trong Chrome
        subprocess.run([chrome_path, file_path])

    def savechart(self):
        # Mở hộp thoại lưu file
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Lưu Biểu đồ", "", "HTML Files (*.html)")

        if file_path:
            # Lưu biểu đồ Sunburst với tên file đã chọn
            if hasattr(self, 'fig'):  # Kiểm tra xem 'fig' đã được tạo chưa
                self.fig.write_html(file_path)
                QtWidgets.QMessageBox.information(self, 'Thông báo', 'Biểu đồ đã được lưu thành công!')
            else:
                QtWidgets.QMessageBox.warning(self, 'Lỗi', 'Biểu đồ chưa được tạo. Vui lòng tạo biểu đồ trước khi lưu.')


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    main_window = MainWindowEx()
    main_window.show()
    sys.exit(app.exec())
