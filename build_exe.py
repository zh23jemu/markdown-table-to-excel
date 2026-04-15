from pathlib import Path
import shutil
import subprocess


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    python_exe = project_dir / ".venv" / "Scripts" / "python.exe"

    if not python_exe.exists():
        raise SystemExit("未找到项目虚拟环境，请先创建 .venv。")

    for folder_name in ["build", "dist"]:
        folder_path = project_dir / folder_name
        if folder_path.exists():
            shutil.rmtree(folder_path)

    spec_path = project_dir / "markdown_to_excel.spec"
    if spec_path.exists():
        spec_path.unlink()

    command = [
        str(python_exe),
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onefile",
        "--hidden-import",
        "tkinterdnd2",
        "--name",
        "Markdown表格导出工具",
        str(project_dir / "markdown_to_excel.py"),
    ]
    subprocess.run(command, check=True, cwd=project_dir)

    exe_path = project_dir / "dist" / "Markdown表格导出工具.exe"
    print(f"打包完成：{exe_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
