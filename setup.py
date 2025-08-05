from cx_Freeze import setup, Executable

# 의존성 패키지들
packages = [
    "pdfplumber", "pandas", "openpyxl", "requests", 
    "camelot", "cv2", "numpy", "PIL", "subprocess", 
    "os", "time", "re", "html", "datetime", "tempfile"
]

# 빌드 옵션
build_exe_options = {
    "packages": packages,
    "excludes": ["tkinter"],
}

setup(
    name="PDFProcessor",
    version="1.0",
    description="PDF Processing Tool with OCR and Translation",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "test.py",
            target_name="pdf_processor.exe",
            base="Console"
        )
    ]
)
