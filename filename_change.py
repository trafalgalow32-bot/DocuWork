# filename_change.py

import os
import shutil
from ctypes import windll, wintypes, byref

INPUT_FOLDER = r"C:\KHK\DocuWork\input"
OUTPUT_FOLDER = r"C:\KHK\DocuWork\output"

# 제거하고 싶은 접두사들 (필요하면 계속 추가하시면 됩니다)
PREFIXES = ["Screen_Recording_", "Screenshot_"]


def strip_prefix(filename):
    """파일명 서두에 PREFIXES 중 하나가 있으면 제거한 이름을 반환."""
    for prefix in PREFIXES:
        if filename.startswith(prefix):
            return filename[len(prefix):]
    return filename


def make_unique(folder, filename):
    """output 폴더에 같은 이름이 이미 있으면 (1), (2)... 붙여 충돌 방지."""
    name, ext = os.path.splitext(filename)
    candidate = filename
    counter = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{name} ({counter}){ext}"
        counter += 1
    return candidate


def set_creation_time(path, creation_timestamp):
    """Windows API로 파일의 '만든 날짜'를 지정한 시각으로 설정."""
    # Unix 타임스탬프(1970 기준) -> Windows FILETIME(1601 기준, 100ns 단위)
    ft = int((creation_timestamp + 11644473600) * 10_000_000)
    filetime = wintypes.FILETIME(ft & 0xFFFFFFFF, ft >> 32)

    FILE_WRITE_ATTRIBUTES = 0x100
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80

    kernel32 = windll.kernel32
    kernel32.CreateFileW.restype = wintypes.HANDLE

    handle = kernel32.CreateFileW(
        path, FILE_WRITE_ATTRIBUTES, 0, None,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None
    )
    if handle == wintypes.HANDLE(-1).value:  # INVALID_HANDLE_VALUE
        raise OSError(f"파일 핸들 열기 실패: {path}")

    try:
        # 인자 순서: (핸들, 생성시각, 접근시각, 수정시각) — 생성시각만 설정
        kernel32.SetFileTime(handle, byref(filetime), None, None)
    finally:
        kernel32.CloseHandle(handle)


def main():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    renamed_count = 0
    copied_count = 0

    for filename in os.listdir(INPUT_FOLDER):
        src_path = os.path.join(INPUT_FOLDER, filename)

        if not os.path.isfile(src_path):
            continue

        # 원본의 '만든 날짜' 미리 읽어두기 (Windows에서 st_ctime = 생성시각)
        original_ctime = os.stat(src_path).st_ctime

        new_name = strip_prefix(filename)
        new_name = make_unique(OUTPUT_FOLDER, new_name)
        dst_path = os.path.join(OUTPUT_FOLDER, new_name)

        # 내용 + 수정시각 보존 복사
        shutil.copy2(src_path, dst_path)

        # 사본의 '만든 날짜'를 원본 값으로 되돌리기
        try:
            set_creation_time(dst_path, original_ctime)
        except OSError as e:
            print(f"[경고] 생성날짜 설정 실패: {new_name} ({e})")

        copied_count += 1
        if new_name != filename:
            renamed_count += 1
            print(f"[변경] {filename}  ->  {new_name}")
        else:
            print(f"[유지] {filename}")

    print(f"\n완료! 총 {copied_count}개 복사, 그 중 {renamed_count}개 이름 변경됨.")


if __name__ == "__main__":
    main()