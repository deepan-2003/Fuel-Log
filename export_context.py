
import os

# Folder to scan
SOURCE_FOLDER = r"C:\Users\ELCOT\Downloads\files"

# Output text file
OUTPUT_FILE = "code_context.txt"

# File extensions to include
CODE_EXTENSIONS = {
    ".py", ".js", ".jsx", ".ts", ".tsx",
    ".html", ".css", ".scss",
    ".java", ".cpp", ".c", ".h", ".cs",
    ".php", ".rb", ".go", ".rs",
    ".sql", ".sh", ".bat",
    ".json", ".xml", ".yaml", ".yml",
    ".md", ".txt", ".ipynb",
    ".dart", ".swift", ".kt",
    ".vue", ".svelte"
}

# Folders to ignore
IGNORE_DIRS = {
    ".git", "__pycache__", "node_modules",
    ".venv", "venv", "dist", "build",
    ".idea", ".vscode"
}

def export_code_context():
    with open(OUTPUT_FILE, "w", encoding="utf-8") as output:

        for root, dirs, files in os.walk(SOURCE_FOLDER):

            # Skip unwanted folders
            dirs[:] = [
                d for d in dirs if d not in IGNORE_DIRS
            ]

            for filename in sorted(files):
                ext = os.path.splitext(filename)[1].lower()

                if ext not in CODE_EXTENSIONS:
                    continue

                file_path = os.path.join(root, filename)
                relative_path = os.path.relpath(
                    file_path, SOURCE_FOLDER
                )

                try:
                    with open(
                        file_path, "r", encoding="utf-8"
                    ) as file:
                        content = file.read()

                    output.write("\n")
                    output.write("=" * 80 + "\n")
                    output.write(f"FILE: {relative_path}\n")
                    output.write("=" * 80 + "\n\n")
                    output.write(content)
                    output.write("\n\n")

                    print(f"Exported: {relative_path}")

                except Exception as e:
                    print(f"Skipped: {relative_path} - {e}")

    print(f"\nExport completed: {OUTPUT_FILE}")


if __name__ == "__main__":
    export_code_context()
