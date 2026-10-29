from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import openpyxl
from PIL import Image
import io
import os
import zipfile
import shutil
import re

app = FastAPI(title="Teacher Utility API", version="2.1")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "temp_storage"
os.makedirs(TEMP_DIR, exist_ok=True)

# --- HÀM HỖ TRỢ XỬ LÝ ĐỀ THI BẢO TOÀN CÔNG THỨC VÀ ĐỊNH DẠNG ---
def parse_docx_to_questions(doc_path):
    doc = Document(doc_path)
    # Cấu trúc lưu trữ câu hỏi và các đoạn văn XML liên quan để giữ công thức toán học (OMML)
    parts = {"part1": [], "part2": []}
    current_part = None
    
    current_q = None
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text and not p._p.xpath('.//*[local-name()="oMath"]'):
            continue
            
        if "phần i" in text.lower():
            current_part = "part1"
            continue
        elif "phần ii" in text.lower():
            current_part = "part2"
            continue
            
        # Nhận diện câu hỏi (Ví dụ: "Câu 1:", "Câu 1.", ...)
        if re.match(r"^câu\s+\d+[\.:]", text.lower()):
            if current_q:
                if current_part == "part1":
                    parts["part1"].append(current_q)
                elif current_part == "part2":
                    parts["part2"].append(current_q)
            
            current_q = {
                "title_p": p, # Giữ nguyên paragraph gốc chứa công thức toán/định dạng
                "options": [],
                "text": text
            }
        elif current_q is not None:
            # Nhận diện đáp án Phần I (A., B., C., D. hoặc *A., *B.,...)
            # Hoặc Phần II (a), b), c), d) hoặc *a), *b),...)
            is_opt_p1 = re.match(r"^(\*)?[a-dđA-DĐ][\.\)]", text)
            if is_opt_p1 or current_part:
                is_correct = text.startswith("*")
                clean_text = text.lstrip("*").strip()
                current_q["options"].append({
                    "p": p,
                    "text": clean_text,
                    "is_correct": is_correct,
                    "raw": text
                })
                
    if current_q:
        if current_part == "part1":
            parts["part1"].append(current_q)
        elif current_part == "part2":
            parts["part2"].append(current_q)
            
    return doc, parts

@app.post("/api/mix-exams")
async def mix_exams(
    file: UploadFile = File(...),
    code_count: int = Form(4),
    start_code: int = Form(101)
):
    import random
    file_path = os.path.join(TEMP_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        doc, parts = parse_docx_to_questions(file_path)
        output_folder = os.path.join(TEMP_DIR, f"output_{start_code}")
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
        os.makedirs(output_folder, exist_ok=True)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Đáp án tổng hợp"
        ws.append(["Mã đề", "Loại", "Câu", "Đáp án đúng"])
        
        generated_files = []
        
        for i in range(code_count):
            curr_code = start_code + i
            new_doc = Document()
            
            # Copy style/spacing cơ bản từ doc gốc nếu muốn hoặc tạo mới
            answer_record = []
            
            # Trộn Phần I
            p1_questions = parts["part1"].copy()
            random.shuffle(p1_questions)
            
            if p1_questions:
                p_head = new_doc.add_paragraph()
                p_head.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.").bold = True
                
            for idx, q in enumerate(p1_questions, 1):
                # Thêm tiêu đề câu hỏi (giữ nguyên XML/công thức nếu có)
                new_p = new_doc.add_paragraph()
                # Clone các run và oMath từ paragraph gốc
                for child in q["title_p"]._p:
                    new_p._p.append(child)
                # Cập nhật lại số thứ tự câu nếu cần, hoặc giữ nguyên cấu trúc
                
                opts = q["options"].copy()
                random.shuffle(opts)
                
                labels = ["A", "B", "C", "D"]
                correct_label = ""
                
                for o_idx, opt in enumerate(opts):
                    lbl = labels[o_idx]
                    if opt["is_correct"]:
                        correct_label = lbl
                        
                    opt_p = new_doc.add_paragraph()
                    opt_p.paragraph_format.left_indent = docx.shared.Pt(18) if 'docx' in globals() else 18
                    run = opt_p.add_run(f"{lbl}. ")
                    run.bold = True
                    for child in opt["p"]._p:
                        # Tránh lặp lại nhãn cũ nếu có trong text, copy XML nội dung
                        opt_p._p.append(child)
                        
                answer_record.append(("Phần I", idx, correct_label))
                
            # Trộn Phần II (Đúng/Sai)
            p2_questions = parts["part2"].copy()
            random.shuffle(p2_questions)
            
            if p2_questions:
                new_doc.add_paragraph().add_run("PHẦN II. Câu trắc nghiệm đúng sai.").bold = True
                
            for idx, q in enumerate(p2_questions, 1):
                new_p = new_doc.add_paragraph()
                for child in q["title_p"]._p:
                    new_p._p.append(child)
                    
                opts = q["options"].copy()
                random.shuffle(opts)
                
                sub_labels = ["a", "b", "c", "d"]
                q_answers = []
                
                for o_idx, opt in enumerate(opts):
                    s_lbl = sub_labels[o_idx]
                    is_true = "Đúng" if opt["is_correct"] else "Sai"
                    q_answers.append(f"{s_lbl}): {is_true}")
                    
                    sub_p = new_doc.add_paragraph()
                    run = sub_p.add_run(f"{s_lbl}) ")
                    run.bold = True
                    for child in opt["p"]._p:
                        sub_p._p.append(child)
                        
                answer_record.append(("Phần II", idx, ", ".join(q_answers)))
                
            # Lưu file docx mã đề
            exam_filename = f"Ma_de_{curr_code}.docx"
            exam_path = os.path.join(output_folder, exam_filename)
            new_doc.save(exam_path)
            generated_files.append(exam_path)
            
            # Ghi vào excel
            for part_name, q_num, ans in answer_record:
                ws.append([curr_code, part_name, q_num, ans])
                
        excel_path = os.path.join(output_folder, "Bang_Dap_An_Tong_Hop.xlsx")
        wb.save(excel_path)
        
        # Nén thành file ZIP
        zip_filename = f"Bo_De_Thi_{start_code}.zip"
        zip_path = os.path.join(TEMP_DIR, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in generated_files:
                zipf.write(f, os.path.basename(f))
            zipf.write(excel_path, os.path.basename(excel_path))
            
        return FileResponse(zip_path, media_type="application/x-zip-compressed", filename=zip_filename)
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý trộn đề: {str(e)}")

# --- API CONVERT ẢNH ---
import docx
@app.post("/api/convert-image")
async def convert_image(file: UploadFile = File(...), target_format: str = Form(...)):
    try:
        image_data = await file.read()
        image = Image.open(io.BytesIO(image_data))
        if target_format.upper() == "JPEG" and image.mode in ("RGBA", "P"):
            image = image.convert("RGB")
            
        orig_name = os.path.splitext(file.filename)[0]
        out_name = f"{orig_name}_converted.{target_format.lower()}"
        out_path = os.path.join(TEMP_DIR, out_name)
        image.save(out_path, format=target_format.upper())
        return FileResponse(out_path, filename=out_name)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- API CONVERT VĂN BẢN (MỚI) ---
@app.post("/api/convert-doc")
async def convert_doc(file: UploadFile = File(...), mode: str = Form(...)):
    try:
        file_data = await file.read()
        orig_name = os.path.splitext(file.filename)[0]
        input_path = os.path.join(TEMP_DIR, file.filename)
        with open(input_path, "wb") as f:
            f.write(file_data)
            
        if mode == "docx_to_txt":
            doc = Document(input_path)
            txt_content = "\n".join([p.text for p in doc.paragraphs])
            out_name = f"{orig_name}.txt"
            out_path = os.path.join(TEMP_DIR, out_name)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(txt_content)
            return FileResponse(out_path, filename=out_name)
        else:
            # Mặc định trả về file text hoặc giả lập convert
            out_name = f"{orig_name}_processed.txt"
            out_path = os.path.join(TEMP_DIR, out_name)
            shutil.copy(input_path, out_path)
            return FileResponse(out_path, filename=out_name)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))