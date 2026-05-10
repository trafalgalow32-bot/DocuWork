# hwpTopdf.py

import win32com.client
import os
import glob

FOLDER = r"C:\KHK\DocuWork\object"

hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

hwp_files = glob.glob(os.path.join(FOLDER, "*.hwp"))
print(f"총 {len(hwp_files)}개 파일 변환 시작...\n")

success, fail = 0, 0

for i, file in enumerate(hwp_files, 1):
    try:
        abs_path = os.path.abspath(file)
        pdf_path = abs_path.replace(".hwp", ".pdf")

        hwp.Open(abs_path, "HWP", "")       # ← 3개 파라미터
        hwp.SaveAs(pdf_path, "PDF", "")     # ← 3개 파라미터

        print(f"[{i}/{len(hwp_files)}] ✅ 완료: {os.path.basename(pdf_path)}")
        success += 1
    except Exception as e:
        print(f"[{i}/{len(hwp_files)}] ❌ 실패: {os.path.basename(file)} → {e}")
        fail += 1

hwp.Quit()
print(f"\n변환 완료! 성공 {success}개 / 실패 {fail}개")