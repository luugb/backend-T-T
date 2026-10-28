import os
import shutil
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/test-single-doc")
async def test_single_doc(file: UploadFile = File(...)):
    temp_input_path = f"temp_{file.filename}"
    output_path = "test_output.docx"
    try:
        # 1. Lưu file tạm
        with open(temp_input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # 2. Tạo một file Word hoàn toàn mới bằng python-docx
        doc = Document()
        doc.add_heading("Bản Test Đề Thi", level=1)
        
        # Đọc file gốc và copy thô tất cả các đoạn văn tìm được vào file mới
        orig_doc = Document(temp_input_path)
        for p in orig_doc.paragraphs:
            if p.text.strip():
                doc.add_paragraph(p.text)

        doc.save(output_path)

        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)

        # Trả về trực tiếp 1 file docx duy nhất để kiểm tra
        return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="Test_Output.docx")

    except Exception as e:
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)
        raise HTTPException(status_code=500, detail=str(e))