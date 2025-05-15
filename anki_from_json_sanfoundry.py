import json
import genanki
import random
import html

# Hàm để tải dữ liệu JSON từ tệp
def load_json_from_file(filename="/home/mintori/Desktop/1000 C# MCQs.json"):
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy tệp '{filename}'. Vui lòng tạo tệp này trong cùng thư mục với script.")
        return None
    except json.JSONDecodeError as e:
        print(f"Lỗi giải mã JSON từ tệp '{filename}': {e}")
        return None
    except Exception as e:
        print(f"Đã xảy ra lỗi không xác định khi đọc tệp '{filename}': {e}")
        return None

# Tải dữ liệu JSON
data = load_json_from_file()

if data is None:
    exit()

# Tạo một model cho Anki
model_id = random.randrange(1 << 30, 1 << 31)
my_model = genanki.Model(
    model_id,
    'Simple Model with Options and Explanation v3', # Cập nhật tên model
    fields=[
        {'name': 'Question'},
        {'name': 'Options'},
        {'name': 'Answer'},
        {'name': 'Explanation'},
        {'name': 'SourceID'},
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}<hr><div style="text-align: left; margin-left: 20px;"><ul>{{Options}}</ul></div><div style="font-size: 0.8em; color: grey; margin-top: 10px; text-align: right;">ID: {{SourceID}}</div>',
            'afmt': '{{FrontSide}}<hr id="answer"><b>{{Answer}}</b><br><br><div style="text-align: left;"><em>{{Explanation}}</em></div>',
        },
    ],
    css="""
    .card {
        font-family: arial;
        font-size: 20px;
        text-align: center; /* Giữ lại center cho card tổng thể, các phần tử con có thể override */
        color: black;
        background-color: white;
    }
    ul {
        list-style-type: none;
        padding-left: 0; /* Bỏ padding mặc định của ul */
        margin: 0;
    }
    li {
        margin: 8px 0;
        text-align: left; /* Căn lề trái cho các mục lựa chọn */
    }
    pre {
        background-color: #f5f5f5;
        border: 1px solid #ccc;
        padding: 10px;
        border-radius: 4px;
        white-space: pre-wrap;
        word-wrap: break-word;
        text-align: left; /* Căn lề trái cho khối code */
    }
    .csharp ol, .csharp pre { /* CSS cho khối code C# từ JSON của bạn */
        text-align: left;
        margin-left: 0;
        padding-left: 1.5em; /* Hoặc một giá trị phù hợp để ol hiển thị số */
    }
    .csharp .de1 { /* Ví dụ style cho class .de1 trong code C# */
        font-family: 'Courier New', Courier, monospace;
    }
    """
)

# Tạo một deck mới
deck_id = random.randrange(1 << 30, 1 << 31)
my_deck = genanki.Deck(
    deck_id,
    'C#/.NET MCQs from JSON File v2' # Cập nhật tên deck
)

# Duyệt qua dữ liệu JSON và tạo notes
for item in data:
    question_text = item.get("Question", "").strip()
    source_id = str(item.get("id", ""))

    if not question_text and source_id == "7": # ID là chuỗi sau khi lấy từ get
        question_text = "[Câu hỏi bị trống - Vui lòng cập nhật]"
        print(f"Mục với id {source_id} có câu hỏi trống. Đã thêm placeholder.")
    elif not question_text:
        print(f"Bỏ qua mục với id {source_id} do câu hỏi trống.")
        continue

    options_list = item.get("Options", [])
    options_html = ""
    if options_list:
        for opt in options_list:
            # Escape HTML trong nội dung của lựa chọn TRƯỚC KHI đưa vào <li>
            escaped_opt = html.escape(str(opt)) # Chuyển opt sang str phòng trường hợp nó là số
            options_html += f"<li>{escaped_opt}</li>"

    answer_full_text = item.get("Answer", "")
    answer_full_text = answer_full_text.replace("\\n", "\n") # Chuẩn hóa newlines

    raw_answer_text_component = ""
    raw_explanation_text_component = ""

    explanation_marker = "\nExplanation: "
    if explanation_marker in answer_full_text:
        parts = answer_full_text.split(explanation_marker, 1)
        raw_answer_text_component = parts[0].replace(" Answer: ", "").strip()
        raw_explanation_text_component = parts[1].strip()
    elif "Explanation: " in answer_full_text:
        explanation_marker_no_nl = "Explanation: "
        parts = answer_full_text.split(explanation_marker_no_nl, 1)
        raw_answer_text_component = parts[0].replace(" Answer: ", "").strip()
        raw_explanation_text_component = parts[1].strip()
    else:
        raw_answer_text_component = answer_full_text.replace(" Answer: ", "").strip()

    # Escape HTML cho phần câu trả lời
    answer_part = html.escape(raw_answer_text_component)

    # Escape HTML cho phần giải thích, SAU ĐÓ thay thế \n bằng <br>
    if raw_explanation_text_component:
        escaped_explanation_text = html.escape(raw_explanation_text_component)
        explanation_part = escaped_explanation_text.replace("\n", "<br>")
    else:
        explanation_part = ""

    # question_text không được escape ở đây vì nó có thể chứa HTML hợp lệ (như câu ID 4)

    my_note = genanki.Note(
        model=my_model,
        fields=[
            question_text,
            options_html,
            f"Đáp án: {answer_part}",
            explanation_part,
            source_id
        ],
        guid=genanki.guid_for(source_id if source_id else str(random.randrange(1 << 30, 1 << 31)))
    )
    my_deck.add_note(my_note)

# Tạo package .apkg
output_filename = 'CSharp_MCQs.apkg'
try:
    genanki.Package(my_deck).write_to_file(output_filename)
    print(f"Tệp '{output_filename}' đã được tạo thành công!")
except Exception as e:
    print(f"Đã xảy ra lỗi khi ghi tệp .apkg: {e}")
