from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware   # <-- add this
from app.analysis import run_analysis_pipeline
import tempfile

app = FastAPI(title="Burton & Bamber Excel Dashboard")

# -----------------------------
#  CORS configuration
# -----------------------------
origins = [
    "https://v0-baile-ai-ai-analyst.vercel.app"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,      # list of allowed origins
    allow_credentials=True,
    allow_methods=["*"],        # allow all HTTP methods (GET, POST, etc.)
    allow_headers=["*"],        # allow all headers
)

# -----------------------------
#  Routes
# -----------------------------
@app.post("/upload/")
async def upload_excel(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    output_path = run_analysis_pipeline(tmp_path)
    return FileResponse(output_path, filename="Burton_Bamber_Dashboard.xlsx")

@app.get("/")
def home():
    return {"message": "Upload Excel files at /upload/ to generate the dashboard"}

# -----------------------------
#  Run server (for local testing)
# -----------------------------
import uvicorn

if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=8000)
