import re
from datetime import datetime, timedelta

def process_schedule(day_data, week_data, subject,teacher_data, class_id):
    # Tách chuỗi
    days = re.findall(r'\s*(\d+):\s*(\d+)-(\d+),([A-Za-z0-9]+)', day_data)
    weeks = re.findall(r'(\d+)-(\d+)', week_data)

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

    # Hàm tính ngày tháng
    def calculate_date(week_num, day_of_week):
        start_date = datetime(2025, 1, 6)  # Tuần 24 bắt đầu từ ngày 01/06/2025 (mm/dd/yyyy)
        days_to_add = (day_of_week - 1) + (week_num - 24) * 7
        return (start_date + timedelta(days=days_to_add)).strftime("%m/%d/%Y")

    # Lưu kết quả sau khi xử lý
    result = []

    # Duyệt qua từng match và chuyển start, end thành thời gian
    for week in weeks:
        start_week, end_week = map(int, week)  # Chuyển start_week và end_week thành số nguyên
        for match in days:
            thu, start, end, location = match
            thu = int(thu)
            start_time = start_class[int(start)]  # Tra cứu thời gian bắt đầu
            end_time = end_class[int(end)]        # Tra cứu thời gian kết thúc
            result.append({
                "class_id" : class_id,
                "teacher_data" : teacher_data,
                "start_date" : calculate_date(start_week,thu),
                "end_date" : calculate_date(end_week,thu),
                "subject": subject,
                "start_week": start_week,
                "end_week": end_week,
                "thu": thu,
                "start_time": start_time,
                "end_time": end_time,
                "room": location
            })

    return result

# Gọi hàm với dữ liệu mẫu
day_data = "Thứ 2: 7-8,C128;3: 7-8,C128;5: 7-8,C128"
week_data = "24-25;29-42"
subject_data = "Tiếng Nhật 4 (CNTT)"
teacher_data = "Phạm Công Thắng"
id_class = "1023220.2420.23.99"

result = process_schedule(day_data, week_data, subject_data, teacher_data, id_class)
# In kết quả
# for item in result:
#     print(f"Môn học: {item['subject']} ngày {item['start_date']}-{item['end_date']}: "
#             f"Thứ {item['thu']}: {item['start_time']} đến {item['end_time']}, Phòng {item['room']}. {item['teacher_data']} {item['class_id']}")


from icalendar import Calendar, Event
from icalendar.prop import vRecur, vDatetime
from datetime import datetime, timedelta
import pytz

import os
import pandas as pd
def excute(path):
    data = pd.read_excel('../book.xlsx')

    result = []
    rows = data.values.tolist()
    for row in rows:
        day_data = row[5]
        week_data = row[6]
        subject_data = row[2]
        teacher_data = row[4]
        id_class = row[1]
        result.extend(process_schedule(day_data, week_data, subject_data, teacher_data, id_class))

    def calculate_date(week_num, day_of_week):
        start_date = datetime(2025, 1, 6)
        days_to_add = (day_of_week - 1) + (week_num - 24) * 7
        return (start_date + timedelta(days=days_to_add))

    cal = Calendar()

    daily = {
        2: "MO",  # Thứ Hai
        3: "TU",  # Thứ Ba
        4: "WE",  # Thứ Tư
        5: "TH",  # Thứ Năm
        6: "FR",  # Thứ Sáu
        7: "SA",  # Thứ Bảy
        8: "SU"   # Chủ Nhật
    }

    timezone = pytz.timezone('Asia/Ho_Chi_Minh')

    for item in result:
        event = Event()

        start_day = calculate_date(item['start_week'], item['thu']).strftime("%Y%m%d") + 'T' + item['start_time'].replace(":", "")
        end_date_str = calculate_date(item['end_week'], item['thu']).strftime("%Y%m%d") + 'T' + item['end_time'].replace(":", "")
        end_day = calculate_date(item['start_week'], item['thu']).strftime("%Y%m%d") + 'T' + item['end_time'].replace(":", "")

        start_day = datetime.strptime(start_day, "%Y%m%dT%H%M%S")
        end_day = datetime.strptime(end_day, "%Y%m%dT%H%M%S")
        end_datetime = datetime.strptime(end_date_str, "%Y%m%dT%H%M%S")

        start_day = timezone.localize(start_day)
        end_day = timezone.localize(end_day)
        end_datetime = timezone.localize(end_datetime)

        event.add('summary', item['subject'])
        event.add('dtstart', start_day)
        event.add('dtend', end_day)
        event.add('location', item['room'])
        event.add('description', f"Giáo viên: {item['teacher_data']}")
        event.add('uid',
                f"{item['class_id']}-{item['thu']}-{item['start_week']}@dkdut")
        event.add('status', 'CONFIRMED')

        end_date = calculate_date(item['end_week'], item['thu'])  # Tính ngày kết thúc

        rrule = vRecur({
            'FREQ': 'WEEKLY',
            'WKST': daily[item['thu']],
            'UNTIL': end_datetime
        })

        event.add('rrule', rrule)

        cal.add_component(event)

    with open(f'{path}\student_dut.ics', 'wb') as f:
        f.write(cal.to_ical())

    print("Đã tạo file ICS thành công!")
from PyQt6.QtWidgets import QApplication, QLabel, QMessageBox, QVBoxLayout, QWidget
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtCore import Qt
import os
class FileDropWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Kéo và Thả File")
        self.setGeometry(100, 100, 400, 200)

        # Tạo giao diện hiển thị
        layout = QVBoxLayout()
        self.label = QLabel("Kéo file .xlsx vào đây", self)
        self.label.setStyleSheet("font-size: 16px; color: blue;")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        self.setLayout(layout)

        # Cho phép chấp nhận thả file
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """Xử lý sự kiện kéo vào."""
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        """Xử lý sự kiện thả file."""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(".xlsx"):
                file_dir = os.path.dirname(file_path)
                success_file = os.path.join(file_dir, "successfully.txt")

                try:
                    # Gọi hàm xử lý chính
                    excute(file_dir)
                    QMessageBox.information(self, "Thông báo", "Finish")
                except Exception as e:
                    QMessageBox.critical(self, "Lỗi", f"Không thể tạo file: {e}")
            else:
                QMessageBox.critical(self, "Lỗi", "File không hợp lệ! Vui lòng thả file .xlsx")


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    main_window = FileDropWidget()
    main_window.show()
    sys.exit(app.exec())

