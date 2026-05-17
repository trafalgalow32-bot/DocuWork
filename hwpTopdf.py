import win32com.client
import os
import glob

INPUT_FOLDER = r"C:\KHK\DocuWork\input"
OUTPUT_FOLDER = r"C:\KHK\DocuWork\output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

hwp_files = (
    glob.glob(os.path.join(INPUT_FOLDER, "*.hwp")) +
    glob.glob(os.path.join(INPUT_FOLDER, "*.hwpx"))
)

print(f"총 {len(hwp_files)}개 파일 변환 시작...\n")

success, fail = 0, 0

for i, file in enumerate(hwp_files, 1):
    try:
        abs_path = os.path.abspath(file)

        filename = os.path.splitext(os.path.basename(file))[0]
        ext = os.path.splitext(file)[1].lower()

        pdf_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, f"{filename}.pdf")
        )

        # 확장자별 포맷 지정
        if ext == ".hwpx":
            hwp.Open(abs_path, "HWPX", "")
        else:
            hwp.Open(abs_path, "HWP", "")

        hwp.SaveAs(pdf_path, "PDF", "")

        print(f"[{i}/{len(hwp_files)}] ✅ 완료: {pdf_path}")
        success += 1

    except Exception as e:
        print(f"[{i}/{len(hwp_files)}] ❌ 실패: {file}")
        print("오류 내용:", e)
        fail += 1

hwp.Quit()

print(f"\n변환 완료! 성공 {success}개 / 실패 {fail}개")