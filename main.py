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

def extract_questions_clean(doc):
    questions = []
    full_text_blocks = []

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

    full_text = "\n".join(full_text_blocks)
    parts = re.split(r'(Câu\s+\d+[\.:])', full_text)

    print(f"DEBUG: Tổng số phần cắt được từ văn bản: {len(parts)}")

    for i in range(1, len(parts), 2):
        q_content = parts[i+1]
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
                questions.append({
                    "question_text": q_text,
                    "options": options[:4]
                })

    print(f"DEBUG: ĐÃ BÓC TÁCH THÀNH CÔNG {len(questions)} CÂU HỎI!")
    return questions

@app.post("/api/mix-exams")
async def mix_exams(
    file: UploadFile = File(...),
    code_count: int = Form(4),
    start_code: int = Form(101)
):
    print(">>> ĐÃ NHẬN ĐƯỢC REQUEST TRỘN ĐỀ TỪ WEB! <<<")
    temp_input_path = f"temp_{file.filename}"
    try:
        with open(temp_input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_input_path)
        questions_data = extract_questions_clean(doc)

        if len(questions_data) == 0:
            raise ValueError("Không tìm thấy cấu trúc câu hỏi!")

        output_dir = "output_exams"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        excel_answer_data = []
        generated_files = []

        for c in range(code_count):
            exam_code = start_code + c
            new_doc = Document()

            if len(doc.sections) > 0:
                src_sec = doc.sections[0]
                new_sec = new_doc.sections[0]
                for p in src_sec.header.paragraphs:
                    if p.text.strip():
                        new_sec.header.add_paragraph(p.text)

            p_title = new_doc.add_paragraph()
            r_title = p_title.add_run(f"MÃ ĐỀ THI: {exam_code}")
            r_title.bold = True
            r_title.font.name = 'Times New Roman'
            r_title.font.size = Pt(13)
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_doc.add_paragraph()

            # Xáo trộn câu hỏi
            shuffled_questions = questions_data.copy()
            random.shuffle(shuffled_questions)

            for idx, q in enumerate(shuffled_questions):
                p_q = new_doc.add_paragraph()
                r_q = p_q.add_run(f"Câu {idx + 1}: {q['question_text']}")
                r_q.bold = True
                r_q.font.name = 'Times New Roman'
                r_q.font.size = Pt(11)

                # Xáo trộn đáp án
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
                    "Câu hỏi": idx + 1,
                    "Đáp án đúng": correct_label
                })
                
                new_doc.add_paragraph()

            exam_file_path = os.path.join(output_dir, f"De_Thi_Ma_{exam_code}.docx")
            new_doc.save(exam_file_path)
            generated_files.append(exam_file_path)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bang_D_p_An"
        ws.append(["Mã đề", "Câu hỏi", "Đáp án đúng"])
        
        for row in excel_answer_data:
            ws.append([row["Mã đề"], row["Câu hỏi"], row["Đáp án đúng"]])
            
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

        print(">>> TẠO FILE VÀ NÉN ZIP THÀNH CÔNG! <<<")
        return FileResponse(zip_path, media_type="application/zip", filename="Bo_De_Thi_Tron.zip")

    except Exception as e:
        print(f"!!! LỖI XẢY RA: {str(e)} !!!")
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)
        raise HTTPException(status_code=500, detail=str(e))