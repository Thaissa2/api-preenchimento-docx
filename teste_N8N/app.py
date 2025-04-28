from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from docxtpl import DocxTemplate
import io
import json

app = FastAPI()

@app.post("/preencher-docx/")
async def preencher_docx(
    file: UploadFile = File(...),
    data: str = Form(...)
):
    try:
        doc = DocxTemplate(io.BytesIO(await file.read()))
        context = json.loads(data)
        doc.render(context)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=documento_preenchido.docx"}
        )

    except Exception as e:
        return {"error": str(e)}
