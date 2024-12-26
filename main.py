import io
import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile

app = FastAPI()

@app.get("/")
def hello_root():
    return {
        "message": "Hello World"
    }

@app.post("/uploadFile/")
async def create_upload_file(file: UploadFile):
    if file.filename.endswith('.xlsx'):
        f = await file.read()
        xlsx = io.BytesIO(f)
        wb = openpyxl.load_workbook(xlsx)
        ws = wb['Sheet1']

        for cells in ws.iter_rows():
            print([cell.value for cell in cells])

        return True
    
    else: 
        raise HTTPException(status_code=400, detail="File must be in xlsx format")
