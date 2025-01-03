import os
import pandas as pd
from datetime import datetime, timedelta
from icalendar import Calendar, Event
from icalendar.prop import vRecur
import pytz
from PyQt6.QtWidgets import QApplication, QLabel, QMessageBox, QVBoxLayout, QWidget
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtCore import Qt
import re


# Hàm xử lý dữ liệu
def process_schedule_optimized(data):
    result = []

    # Thời gian bắt đầu và kết thúc cho mỗi tiết
    start_class = {
        1: "7:00:00", 2: "8:00:00", 3: "9:00:00", 4: "10:00:00",
        5: "11:00:00", 6: "12:30:00", 7: "13:30:00", 8: "14:30:00",
        9: "15:30:00", 10: "16:30:00"
    }
    end_class = {
        1: "7:50:00", 2: "8:50:00", 3: "9:50:00", 4: "10:50:00",
        5: "11:50:00", 6: "13:20:00", 7: "14:20:00", 8: "15:20:00",
        9: "16:20:00", 10: "17:20:00"
    }

    # Hàm tính toán ngày bắt đầu dựa trên tuần và thứ
    def calculate_date(week_num, day_of_week):
        start_date = datetime(2025, 1, 5)  # Tuần 24 bắt đầu từ ngày 05/01/2025
        days_to_add = (day_of_week - 1) + (week_num - 24) * 7
        return start_date + timedelta(days=days_to_add)

    for index, row in data.iterrows():
        day_data = row['DayData']
        week_data = row['WeekData']
        subject_data = row['Subject']
        teacher_data = row['Teacher']
        id_class = row['ClassID']

        # Tách chuỗi ngày và tuần
        days = re.findall(r'\s*(\d+):\s*(\d+)-(\d+),([A-Za-z0-9]+)', day_data)
        weeks = re.findall(r'(\d+)-(\d+)', week_data)

        for week in weeks:
            start_week, end_week = map(int, week)
            for match in days:
                thu, start, end, location = match
                thu = int(thu)
                start_time = start_class[int(start)]
                end_time = end_class[int(end)]
                result.append({
                    "class_id": id_class,
                    "teacher_data": teacher_data,
                    "start_date": calculate_date(start_week, thu),
                    "end_date": calculate_date(end_week, thu),
                    "subject": subject_data,
                    "start_week": start_week,
                    "end_week": end_week,
                    "thu": thu,
                    "start_time": start_time,
                    "end_time": end_time,
                    "room": location
                })

    return result


# Hàm chính xử lý file Excel và tạo file ICS
def excute_optimized(file_path):
    try:
        # Đọc dữ liệu từ file Excel
        data = pd.read_excel(file_path, usecols=["Mã lớp học phần", "Tên lớp học phần", "Giảng viên", "Thời khóa biểu",
                                                 "Tuần học"])

        # Đổi tên cột
        data.rename(columns={
            "Mã lớp học phần": "ClassID",
            "Tên lớp học phần": "Subject",
            "Giảng viên": "Teacher",
            "Thời khóa biểu": "DayData",
            "Tuần học": "WeekData"
        }, inplace=True)

        # Gọi hàm xử lý
        result = process_schedule_optimized(data)

        # Tạo file ICS
        cal = Calendar()
        timezone = pytz.timezone('Asia/Ho_Chi_Minh')
        daily = {2: "MO", 3: "TU", 4: "WE", 5: "TH", 6: "FR", 7: "SA", 8: "SU"}

        for item in result:
            event = Event()
            start_day = item['start_date'].strftime("%Y%m%d") + 'T' + item['start_time'].replace(":", "")
            end_day = item['start_date'].strftime("%Y%m%d") + 'T' + item['end_time'].replace(":", "")
            end_date = item['end_date'].strftime("%Y%m%d") + 'T' + item['end_time'].replace(":", "")

            start_day = timezone.localize(datetime.strptime(start_day, "%Y%m%dT%H%M%S"))
            end_day = timezone.localize(datetime.strptime(end_day, "%Y%m%dT%H%M%S"))
            end_datetime = timezone.localize(datetime.strptime(end_date, "%Y%m%dT%H%M%S"))

            event.add('summary', item['subject'])
            event.add('dtstart', start_day)
            event.add('dtend', end_day)
            event.add('location', item['room'])
            event.add('description', f"{item['class_id']} - Giáo viên: {item['teacher_data']}")
            event.add('uid', f"{item['class_id']}-{item['thu']}-{item['start_week']}@dkdut")
            event.add('status', 'CONFIRMED')
            rrule = vRecur({
                'FREQ': 'WEEKLY',
                'WKST': daily[item['thu']],
                'UNTIL': end_datetime
            })
            event.add('rrule', rrule)
            cal.add_component(event)

        output_path = os.path.join(os.path.dirname(file_path), 'student_dut.ics')
        with open(output_path, 'wb') as f:
            f.write(cal.to_ical())

        print("Đã tạo file ICS thành công!")

    except Exception as e:
        raise ValueError(f"Lỗi khi xử lý file: {e}")


# Lớp giao diện kéo thả file
class FileDropWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Kéo và Thả File")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()
        self.label = QLabel("Kéo file .xlsx vào đây", self)
        self.label.setStyleSheet("font-size: 16px; color: blue;")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        self.setLayout(layout)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(".xlsx"):
                try:
                    excute_optimized(file_path)
                    QMessageBox.information(self, "Thông báo", "Tạo file thành công!")
                except Exception as e:
                    QMessageBox.critical(self, "Lỗi", f"Không thể tạo file: {e}")
            else:
                QMessageBox.critical(self, "Lỗi", "File không hợp lệ! Vui lòng thả file .xlsx")


# Chạy ứng dụng
if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    main_window = FileDropWidget()
    main_window.show()
    sys.exit(app.exec())
