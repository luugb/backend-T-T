import os
import shutil
import zipfile
import random
import re
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def parse_exam_structure(doc):
    """
    Bóc tách đề thi chuẩn xác thành Phần I (4 lựa chọn) và Phần II (Đúng/Sai),
    đồng thời giữ lại nội dung đoạn văn của các câu hỏi (bao gồm cả đoạn trích dẫn tư liệu).
    """
    full_text_blocks = []
    
    # Lưu lại toàn bộ các paragraph gốc để có thể copy lại nguyên văn phần tư liệu/nội dung câu hỏi
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            full_text_blocks.append(t)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = p.text.strip()
                    if t:
                        full_text_blocks.append(t)

    ignore_keywords = [
        "bộ giáo dục", "kỳ thi", "môn thi", "thí sinh không được", 
        "giám thị", "đề thi chính thức", "đề thi có", "thời gian làm bài", 
        "hết", "trang", "họ, tên thí sinh", "số báo danh"
    ]

    filtered_blocks = []
    for t in full_text_blocks:
        t_lower = t.lower()
        is_meta = False
        if any(kw in t_lower for kw in ignore_keywords) and len(t) < 130 and not t_lower.startswith("câu"):
            is_meta = True
        if "-------" in t or "===" in t or "___" in t:
            is_meta = True
        
        if not is_meta:
            filtered_blocks.append(t)

    full_text = "\n".join(filtered_blocks)

    part1_text = full_text
    part2_text = ""

    part2_match = re.search(r'(PHẦN\s+(?:II|2)[^\n]*)', full_text, re.IGNORECASE)
    if part2_match:
        split_pos = part2_match.start()
        part1_text = full_text[:split_pos]
        part2_text = full_text[split_pos:]

    # --- BÓC TÁCH PHẦN I ---
    part1_questions = []
    p1_parts = re.split(r'(Câu\s+\d+[\.:])', part1_text)

    for i in range(1, len(p1_parts), 2):
        q_content = p1_parts[i+1]
        option_matches = list(re.finditer(r'(?:^|\n)\s*(\*?\s*[A-D][\.\)])', q_content))

        if len(option_matches) >= 4:
            first_opt_pos = option_matches[0].start()
            q_text = q_content[:first_opt_pos].strip()
            
            if q_text.startswith(":") or q_text.startswith("."):
                q_text = q_text[1:].strip()

            options = []
            for j in range(len(option_matches)):
                start_pos = option_matches[j].start()
                end_pos = option_matches[j+1].start() if j + 1 < len(option_matches) else len(q_content)
                opt_raw = q_content[start_pos:end_pos].strip()

                is_correct = opt_raw.startswith("*")
                clean_opt = opt_raw.lstrip("* ").strip()
                
                if len(clean_opt) > 2 and clean_opt[0].upper() in ["A", "B", "C", "D"] and clean_opt[1] in ['.', ')']:
                    clean_opt = clean_opt[2:].strip()

                options.append({
                    "text": clean_opt,
                    "correct": is_correct
                })

            if len(options) >= 4:
                part1_questions.append({
                    "question_text": q_text,
                    "options": options[:4]
                })

    # --- BÓC TÁCH PHẦN II (ĐÚNG / SAI) ---
    part2_questions = []
    if part2_text:
        p2_parts = re.split(r'(Câu\s+\d+[\.:])', part2_text)

        for i in range(1, len(p2_parts), 2):
            q_content = p2_parts[i+1]
            sub_option_matches = list(re.finditer(r'(?:^|\n)\s*(\*?\s*[a-d][\.\)])', q_content))

            if len(sub_option_matches) >= 4:
                first_opt_pos = sub_option_matches[0].start()
                q_text = q_content[:first_opt_pos].strip()
                
                if q_text.startswith(":") or q_text.startswith("."):
                    q_text = q_text[1:].strip()

                sub_options = []
                for j in range(len(sub_option_matches)):
                    start_pos = sub_option_matches[j].start()
                    end_pos = sub_option_matches[j+1].start() if j + 1 < len(sub_option_matches) else len(q_content)
                    opt_raw = q_content[start_pos:end_pos].strip()

                    is_true = opt_raw.startswith("*")
                    clean_opt = opt_raw.lstrip("* ").strip()
                    
                    if len(clean_opt) > 2 and clean_opt[0].lower() in ["a", "b", "c", "d"] and clean_opt[1] in ['.', ')']:
                        clean_opt = clean_opt[2:].strip()

                    sub_options.append({
                        "text": clean_opt,
                        "is_true": is_true
                    })

                if len(sub_options) >= 4:
                    part2_questions.append({
                        "question_text": q_text,
                        "sub_options": sub_options[:4]
                    })

    return part1_questions, part2_questions

@app.post("/api/mix-exams")
async def mix_exams(
    file: UploadFile = File(...),
    code_count: int = Form(4),
    start_code: int = Form(101)
):
    temp_input_path = f"temp_{file.filename}"
    try:
        with open(temp_input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_input_path)
        p1_questions, p2_questions = parse_exam_structure(doc)

        if len(p1_questions) == 0 and len(p2_questions) == 0:
            raise ValueError("Không tìm thấy cấu trúc câu hỏi hợp lệ trong file đề!")

        output_dir = "output_exams"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        excel_answer_data = []
        generated_files = []

        for c in range(code_count):
            exam_code = start_code + c
            new_doc = Document()

            # Thiết lập lề chuẩn cho văn bản
            for section in new_doc.sections:
                section.top_margin = Inches(0.8)
                section.bottom_margin = Inches(0.8)
                section.left_margin = Inches(0.8)
                section.right_margin = Inches(0.8)

            # Sao chép thông tin phần đầu (Header) từ file gốc sang file mới nếu có
            header_copied = False
            for p in doc.paragraphs:
                t = p.text.strip()
                if "BỘ GIÁO DỤC" in t.upper() or "KỲ THI" in t.upper() or "ĐỀ THI" in t.upper():
                    new_p = new_doc.add_paragraph()
                    new_p.paragraph_format.space_after = Pt(2)
                    new_p.paragraph_format.line_spacing = 1.0
                    r = new_p.add_run(t)
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(10)
                    if p.runs and p.runs[0].bold:
                        r.bold = True
                    header_copied = True
                elif header_copied and ("Họ, tên thí sinh" in t or "Số báo danh" in t):
                    break

            if not header_copied:
                # Fallback header chuẩn nếu file gốc không khớp định dạng text cứng
                p_h = new_doc.add_paragraph()
                p_h.paragraph_format.space_after = Pt(2)
                r = p_h.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO                     KỲ THI TỐT NGHIỆP TRUNG HỌC PHỔ THÔNG\nĐỀ THI CHÍNH THỨC                                       Môn thi: LỊCH SỬ")
                r.bold = True
                r.font.name = 'Times New Roman'
                r.font.size = Pt(10)

            # Dòng điền thông tin thí sinh & Mã đề sát nhau
            p_info = new_doc.add_paragraph()
            p_info.paragraph_format.space_after = Pt(6)
            p_info.paragraph_format.line_spacing = 1.0
            r_info = p_info.add_run(f"Họ, tên thí sinh: ...........................................................................\nSố báo danh: .............................................................  Mã đề: {exam_code}")
            r_info.bold = True
            r_info.font.name = 'Times New Roman'
            r_info.font.size = Pt(10.5)

            # ================= GHI PHẦN I =================
            if len(p1_questions) > 0:
                p_p1 = new_doc.add_paragraph()
                p_p1.paragraph_format.space_after = Pt(4)
                p_p1.paragraph_format.line_spacing = 1.0
                r_p1 = p_p1.add_run("PHẦN I. Thí sinh trả lời từ câu 1 đến câu số lượng. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
                r_p1.bold = True
                r_p1.font.name = 'Times New Roman'
                r_p1.font.size = Pt(11)

                shuffled_p1 = p1_questions.copy()
                random.shuffle(shuffled_p1)

                for idx, q in enumerate(shuffled_p1):
                    p_q = new_doc.add_paragraph()
                    p_q.paragraph_format.space_after = Pt(2)
                    p_q.paragraph_format.line_spacing = 1.0
                    r_q = p_q.add_run(f"Câu {idx + 1}. {q['question_text']}")
                    r_q.bold = True
                    r_q.font.name = 'Times New Roman'
                    r_q.font.size = Pt(11)

                    options = q["options"].copy()
                    random.shuffle(options)

                    labels = ["A", "B", "C", "D"]
                    correct_label = "A"

                    # Kiểm tra nếu các đáp án ngắn thì trình bày 2 cột ngang sát nhau gọn gàng
                    max_opt_len = max([len(opt['text']) for opt in options])
                    is_short_options = max_opt_len < 45

                    if is_short_options:
                        tbl_opt = new_doc.add_table(rows=2, cols=2)
                        tbl_opt.autofit = False
                        tbl_opt.columns[0].width = Inches(3.25)
                        tbl_opt.columns[1].width = Inches(3.25)

                        cell_coords = [(0, 0), (0, 1), (1, 0), (1, 1)]
                        for opt_i, opt in enumerate(options):
                            if opt_i < 4:
                                lbl = labels[opt_i]
                                r_idx, c_idx = cell_coords[opt_i]
                                cell = tbl_opt.cell(r_idx, c_idx)
                                p_cell = cell.paragraphs[0]
                                p_cell.paragraph_format.space_after = Pt(1)
                                p_cell.paragraph_format.line_spacing = 1.0
                                r_cell = p_cell.add_run(f"{lbl}. {opt['text']}")
                                r_cell.font.name = 'Times New Roman'
                                r_cell.font.size = Pt(11)

                                if opt["correct"]:
                                    correct_label = lbl
                    else:
                        for opt_i, opt in enumerate(options):
                            if opt_i < len(labels):
                                lbl = labels[opt_i]
                                p_opt = new_doc.add_paragraph()
                                p_opt.paragraph_format.space_after = Pt(2)
                                p_opt.paragraph_format.line_spacing = 1.0
                                p_opt.paragraph_format.left_indent = Inches(0.2)
                                r_opt = p_opt.add_run(f"{lbl}. {opt['text']}")
                                r_opt.font.name = 'Times New Roman'
                                r_opt.font.size = Pt(11)
                                
                                if opt["correct"]:
                                    correct_label = lbl

                    excel_answer_data.append({
                        "Mã đề": exam_code,
                        "Phần": "Phần I",
                        "Câu hỏi": f"Câu {idx + 1}",
                        "Đáp án đúng": correct_label
                    })

            # ================= GHI PHẦN II =================
            if len(p2_questions) > 0:
                p_p2 = new_doc.add_paragraph()
                p_p2.paragraph_format.space_before = Pt(6)
                p_p2.paragraph_format.space_after = Pt(4)
                p_p2.paragraph_format.line_spacing = 1.0
                r_p2 = p_p2.add_run("PHẦN II. Thí sinh trả lời từ câu 1 đến câu hỏi. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
                r_p2.bold = True
                r_p2.font.name = 'Times New Roman'
                r_p2.font.size = Pt(11)

                shuffled_p2 = p2_questions.copy()
                random.shuffle(shuffled_p2)

                for idx, q in enumerate(shuffled_p2):
                    p_q = new_doc.add_paragraph()
                    p_q.paragraph_format.space_after = Pt(2)
                    p_q.paragraph_format.line_spacing = 1.0
                    r_q = p_q.add_run(f"Câu {idx + 1}. {q['question_text']}")
                    r_q.bold = True
                    r_q.font.name = 'Times New Roman'
                    r_q.font.size = Pt(11)

                    sub_opts = q["sub_options"].copy()
                    random.shuffle(sub_opts)

                    sub_labels = ["a", "b", "c", "d"]

                    for s_i, s_opt in enumerate(sub_opts):
                        if s_i < len(sub_labels):
                            lbl = sub_labels[s_i]
                            p_sub = new_doc.add_paragraph()
                            p_sub.paragraph_format.space_after = Pt(2)
                            p_sub.paragraph_format.line_spacing = 1.0
                            p_sub.paragraph_format.left_indent = Inches(0.2)
                            r_sub = p_sub.add_run(f"{lbl}) {s_opt['text']}")
                            r_sub.font.name = 'Times New Roman'
                            r_sub.font.size = Pt(11)
                            
                            ans_text = "Đúng" if s_opt["is_true"] else "Sai"
                            excel_answer_data.append({
                                "Mã đề": exam_code,
                                "Phần": "Phần II",
                                "Câu hỏi": f"Câu {idx + 1} ({lbl})",
                                "Đáp án đúng": ans_text
                            })

            exam_file_path = os.path.join(output_dir, f"De_Thi_Ma_{exam_code}.docx")
            new_doc.save(exam_file_path)
            generated_files.append(exam_file_path)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bang_Dap_An"
        ws.append(["Mã đề", "Phần", "Câu hỏi", "Đáp án đúng"])
        
        for row in excel_answer_data:
            ws.append([row["Mã đề"], row["Phần"], row["Câu hỏi"], row["Đáp án đúng"]])
            
        excel_path = os.path.join(output_dir, "Tong_Hop_Dap_An_Cac_De.xlsx")
        wb.save(excel_path)
        generated_files.append(excel_path)

        zip_filename = "Bo_De_Thi_Tron.zip"
        zip_path = os.path.join(output_dir, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for fpath in generated_files:
                if not fpath.endswith(zip_filename):
                    zipf.write(fpath, os.path.basename(fpath))

        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)

        return FileResponse(zip_path, media_type="application/zip", filename="Bo_De_Thi_Tron.zip")

    except Exception as e:
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)
        raise HTTPException(status_code=500, detail=str(e))