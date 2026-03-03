import subprocess
import os
import shutil
from pathlib import Path


def get_pandoc_bin():
    """
    自動偵測環境中的 pandoc 執行檔路徑。
    """
    # 優先尋找系統路徑中的 pandoc (這在 active 的 conda 環境中通常能直接找到)
    pandoc_path = shutil.which("pandoc")

    if pandoc_path:
        return pandoc_path

    # 備用方案：如果 shutil.which 找不到，嘗試從環境變數找
    conda_prefix = os.environ.get("CONDA_PREFIX")
    if conda_prefix:
        # Windows conda 環境的標準二進位路徑
        potential_path = Path(conda_prefix) / "Library" / "bin" / "pandoc.exe"
        if potential_path.exists():
            return str(potential_path)

    return "pandoc"  # 最後手段：直接回傳名稱，由 OS 嘗試執行


def batch_convert_docx_to_md(input_folder: str, output_folder: str):
    # 取得當前腳本所在的目錄，確保路徑相對於專案根目錄
    base_dir = Path(__file__).parent.resolve()
    input_path = (base_dir / input_folder).resolve()
    output_path = (base_dir / output_folder).resolve()

    # 取得動態 pandoc 路徑
    pandoc_bin = get_pandoc_bin()
    print(f"使用 Pandoc 引擎路徑: {pandoc_bin}")

    if not input_path.exists():
        print(f"錯誤：找不到輸入目錄 {input_path}")
        return

    output_path.mkdir(parents=True, exist_ok=True)
    docx_files = list(input_path.glob("*.docx"))

    if not docx_files:
        print("資訊：未找到 .docx 檔案")
        return

    print(f"開始處理 {len(docx_files)} 個檔案...")

    for file_path in docx_files:
        file_name = file_path.stem

        # 建立專屬輸出目錄：output/檔名/
        target_dir = output_path / file_name
        target_dir.mkdir(parents=True, exist_ok=True)

        output_md = f"{file_name}.md"

        print(f"正在處理: {file_path.name}")

        cmd = [
            pandoc_bin,
            str(file_path),
            "-t",
            "gfm",
            "-o",
            output_md,
            "--extract-media=.",
            "--wrap=none",
        ]

        try:
            result = subprocess.run(
                cmd,
                cwd=str(target_dir),
                capture_output=True,
                text=True,
                encoding="utf-8",
            )

            if result.returncode == 0:
                print(f"--- 成功：已存至 {target_dir}")
            else:
                print(f"--- Pandoc 錯誤: {result.stderr}")

        except Exception as e:
            print(f"--- 系統錯誤: {str(e)}")


if __name__ == "__main__":
    batch_convert_docx_to_md("input", "output")
