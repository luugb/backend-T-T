import os
import shutil
import zipfile
import random
import re
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt
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
    Phân tách đề thi thành 2 phần:
    - Phần I: Trắc nghiệm 4 lựa chọn (A, B, C, D)
    - Phần II: Trắc nghiệm Đúng/Sai (a, b, c, d)
    """
    full_text_blocks = []

    # Thu thập văn bản từ paragraphs và bảng biểu (2 cột)
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

    # Loại bỏ các thông tin rác / tiêu đề / hướng dẫn phòng thi
    ignore_keywords = [
        "bộ giáo dục", "kỳ thi", "môn thi", "thí sinh không được", 
        "giám thị", "đề thi chính thức", "đề thi có", "thời gian làm bài", 
        "hết", "trang"
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

    # Nhận diện vị trí bắt đầu Phần I và Phần II
    part1_text = full_text
    part2_text = ""

    # Tìm ranh giới Phần II (thường chứa "PHẦN II", "Phần II" hoặc "PHẦN 2")
    part2_match = re.search(r'(PHẦN\s+(?:II|2)[^\n]*)', full_text, re.IGNORECASE)
    if part2_match:
        split_pos = part2_match.start()
        part1_text = full_text[:split_pos]
        part2_text = full_text[split_pos:]

    # --- 1. BÓC TÁCH PHẦN I (TRẮC NGHIỆM 4 LỰA CHỌN) ---
    part1_questions = []
    p1_parts = re.split(r'(Câu\s+\d+[\.:])', part1_text)

    for i in range(1, len(p1_parts), 2):
        q_content = p1_parts[i+1]
        option_matches = list(re.finditer(r'(\*?\s*[A-D][\.\)])', q_content))

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

    # --- 2. BÓC TÁCH PHẦN II (ĐÚNG / SAI) ---
    part2_questions = []
    if part2_text:
        p2_parts = re.split(r'(Câu\s+\d+[\.:])', part2_text)

        for i in range(1, len(p2_parts), 2):
            q_content = p2_parts[i+1]
            # Tìm các ý nhỏ dạng *a), a), *b), b)...
            sub_option_matches = list(re.finditer(r'(\*?\s*[a-d][\.\)])', q_content))

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

                    # Có dấu * ở đầu là ĐÚNG, không có dấu * là SAI
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

            # Giữ lại Header nếu có
            if len(doc.sections) > 0:
                src_sec = doc.sections[0]
                new_sec = new_doc.sections[0]
                for p in src_sec.header.paragraphs:
                    if p.text.strip():
                        new_sec.header.add_paragraph(p.text)

            # In tiêu đề mã đề
            p_title = new_doc.add_paragraph()
            r_title = p_title.add_run(f"MÃ ĐỀ THI: {exam_code}")
            r_title.bold = True
            r_title.font.name = 'Times New Roman'
            r_title.font.size = Pt(13)
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_doc.add_paragraph()

            # ================= GHI PHẦN I =================
            if len(p1_questions) > 0:
                p_p1_head = new_doc.add_paragraph()
                r_p1_head = p_p1_head.add_run("PHẦN I. Câu hỏi trắc nghiệm nhiều phương án lựa chọn")
                r_p1_head.bold = True
                r_p1_head.font.name = 'Times New Roman'
                r_p1_head.font.size = Pt(12)

                shuffled_p1 = p1_questions.copy()
                random.shuffle(shuffled_p1)

                for idx, q in enumerate(shuffled_p1):
                    p_q = new_doc.add_paragraph()
                    r_q = p_q.add_run(f"Câu {idx + 1}: {q['question_text']}")
                    r_q.bold = True
                    r_q.font.name = 'Times New Roman'
                    r_q.font.size = Pt(11)

                    options = q["options"].copy()
                    random.shuffle(options)

                    labels = ["A", "B", "C", "D"]
                    correct_label = "A"

                    for opt_i, opt in enumerate(options):
                        if opt_i < len(labels):
                            lbl = labels[opt_i]
                            p_opt = new_doc.add_paragraph()
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
                    new_doc.add_paragraph()

            # ================= GHI PHẦN II =================
            if len(p2_questions) > 0:
                p_p2_head = new_doc.add_paragraph()
                r_p2_head = p_p2_head.add_run("PHẦN II. Câu hỏi trắc nghiệm Đúng/Sai")
                r_p2_head.bold = True
                r_p2_head.font.name = 'Times New Roman'
                r_p2_head.font.size = Pt(12)

                shuffled_p2 = p2_questions.copy()
                random.shuffle(shuffled_p2)

                for idx, q in enumerate(shuffled_p2):
                    p_q = new_doc.add_paragraph()
                    r_q = p_q.add_run(f"Câu {idx + 1}: {q['question_text']}")
                    r_q.bold = True
                    r_q.font.name = 'Times New Roman'
                    r_q.font.size = Pt(11)

                    sub_opts = q["sub_options"].copy()
                    random.shuffle(sub_opts) # Xáo trộn thứ tự các ý a, b, c, d

                    sub_labels = ["a", "b", "c", "d"]

                    for s_i, s_opt in enumerate(sub_opts):
                        if s_i < len(sub_labels):
                            lbl = sub_labels[s_i]
                            p_sub = new_doc.add_paragraph()
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

                    new_doc.add_paragraph()

            exam_file_path = os.path.join(output_dir, f"De_Thi_Ma_{exam_code}.docx")
            new_doc.save(exam_file_path)
            generated_files.append(exam_file_path)

        # Tạo file Excel tổng hợp đáp án
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bang_Dap_An"
        ws.append(["Mã đề", "Phần", "Câu hỏi", "Đáp án đúng"])
        
        for row in excel_answer_data:
            ws.append([row["Mã đề"], row["Phần"], row["Câu hỏi"], row["Đáp án đúng"]])
            
        excel_path = os.path.join(output_dir, "Tong_Hop_Dap_An_Cac_De.xlsx")
        wb.save(excel_path)
        generated_files.append(excel_path)

        # Nén file ZIP
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