import win32com.client
import os
import glob

INPUT_FOLDER = r"C:\KHK\DocuWork\input"
OUTPUT_FOLDER = r"C:\KHK\DocuWork\output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

# 저장 시 "레이아웃이 틀어질 수 있습니다" 예/아니오 팝업 → 자동 "예"
hwp.SetMessageBoxMode(0x10000)

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
        docx_path = os.path.abspath(os.path.join(OUTPUT_FOLDER, f"{filename}.docx"))

        # 포맷은 "" 로 비워 자동 감지 (hwp/hwpx 모두 안전)
        if not hwp.Open(abs_path, "", "versionwarning:false"):
            raise RuntimeError("Open 실패 (반환값 False)")

        # 핵심: docx 저장 포맷은 "OOXML"
        if not hwp.SaveAs(docx_path, "OOXML", ""):
            raise RuntimeError("SaveAs 실패 (반환값 False)")

        # 실제로 파일이 생겼는지 최종 확인
        if not os.path.exists(docx_path):
            raise RuntimeError("SaveAs는 True인데 파일이 실제로 없음")

        print(f"[{i}/{len(hwp_files)}] ✅ 완료: {docx_path}")
        success += 1

    except Exception as e:
        print(f"[{i}/{len(hwp_files)}] ❌ 실패: {file}")
        print("오류 내용:", e)
        fail += 1

hwp.Quit()
print(f"\n변환 완료! 성공 {success}개 / 실패 {fail}개")