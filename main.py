import os
import shutil
import zipfile
import random
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import openpyxl

app = FastAPI()

# Cho phép Frontend gọi API lên Backend (CORS)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/mix-exams")
async def mix_exams(
    file: UploadFile = File(...),
    code_count: int = Form(4),
    start_code: int = Form(101)
):
    temp_input_path = f"temp_{file.filename}"
    with open(temp_input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Đọc file Word gốc bằng python-docx
    doc = Document(temp_input_path)
    
    # 1. Phân tích tài liệu để tách phần Header (tiêu đề trường, lớp...) và các Câu hỏi
    # (Quy ước: Câu hỏi bắt đầu bằng từ "Câu X:" hoặc "Câu X.")
    paragraphs = doc.paragraphs
    
    # Thuật toán trích xuất câu hỏi từ các đoạn văn bản
    questions_data = []
    current_q = None
    
    for p in paragraphs:
        text = p.text.strip()
        if text.startswith("Câu ") and ("." in text or ":" in text):
            if current_q:
                questions_data.append(current_q)
            current_q = {"title_para": p, "options": []}
        elif current_q and (text.startswith("A.") or text.startswith("B.") or text.startswith("C.") or text.startswith("D.") or text.startswith("*")):
            is_correct = "*" in text
            current_q["options"].append({"para": p, "text": text, "correct": is_correct})
            
    if current_q:
        questions_data.append(current_q)

    output_dir = "output_exams"
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    excel_answer_data = []
    generated_files = []

    # 2. Vòng lặp sinh ra các mã đề theo yêu cầu
    for c in range(code_count):
        exam_code = start_code + c
        
        # Copy lại một bản doc gốc để giữ nguyên toàn bộ định dạng, font, header
        exam_doc = Document(temp_input_path)
        
        # Xáo trộn danh sách câu hỏi
        shuffled_q = questions_data.copy()
        random.shuffle(shuffled_q)
        
        # (Thực hiện ghi đè nội dung đã xáo trộn vào exam_doc...)
        
        exam_file_path = os.path.join(output_dir, f"De_Thi_Ma_{exam_code}.docx")
        exam_doc.save(exam_file_path)
        generated_files.append(exam_file_path)

        # Giả lập ghi dữ liệu đáp án vào Excel
        for idx, q in enumerate(shuffled_q):
            excel_answer_data.append({
                "Mã đề": exam_code,
                "Câu hỏi": idx + 1,
                "Đáp án đúng": "A" # Tạm thời minh họa
            })

    # 3. Tạo file Excel tổng hợp đáp án bằng openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bang_Dap_An"
    ws.append(["Mã đề", "Câu hỏi", "Đáp án đúng"])
    
    for row in excel_answer_data:
        ws.append([row["Mã đề"], row["Câu hỏi"], row["Đáp án đúng"]])
        
    excel_path = os.path.join(output_dir, "Tong_Hop_Dap_An_Cac_De.xlsx")
    wb.save(excel_path)
    generated_files.append(excel_path)

    # 4. Đóng gói toàn bộ file Word và Excel vào file ZIP
    zip_filename = "Bo_De_Thi_Tron.zip"
    zip_path = os.path.join(output_dir, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for fpath in generated_files:
            if not fpath.endswith(zip_filename):
                zipf.write(fpath, os.path.basename(fpath))

    # Xóa file upload tạm
    if os.path.exists(temp_input_path):
        os.remove(temp_input_path)

    return FileResponse(zip_path, media_type="application/zip", filename="Bo_De_Thi_Tron.zip")