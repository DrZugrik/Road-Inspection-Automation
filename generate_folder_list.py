"""
generate_folder_list.py
Генерирует список FOLDERS_TO_PROCESS для всех подпапок в указанной директории.
"""

from pathlib import Path

# ============================================================
# Укажите путь к папке, которую нужно проанализировать
# ============================================================
SOURCE_DIR = r"C:\Users\user\Downloads\4 этап - доделки 01,10,2026"
# ============================================================

def generate_list():
    target_path = Path(SOURCE_DIR)
    
    if not target_path.exists():
        print(f"❌ Папка не найдена: {SOURCE_DIR}")
        return

    # Получаем все подпапки внутри SOURCE_DIR (только папки, не файлы)
    folders = [f for f in target_path.iterdir() if f.is_dir()]
    
    # Сортируем их по имени для наглядности
    folders.sort(key=lambda x: x.name)
    
    if not folders:
        print("⚠️ В указанной папке не найдено подпапок.")
        return

    print("\n" + "=" * 60)
    print(f"📂 Найдено папок: {len(folders)}")
    print("=" * 60)
    print(f"\nFOLDERS_TO_PROCESS = [")
    for f in folders:
        # Используем сырой строковый литерал (r"...") для корректного отображения слешей
        print(f"    r\"{f}\",")
    print("]")
    print("\n✅ Готово! Скопируйте этот список и вставьте его в main.py.\n")

if __name__ == "__main__":
    generate_list()