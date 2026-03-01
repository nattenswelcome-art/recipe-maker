import os
import glob
import subprocess
import json
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRAINING_DIR = os.path.join(BASE_DIR, 'training_data')
OUTPUT_JSON = os.path.join(BASE_DIR, 'training_dataset.json')

def parse_docx_raw(docx_path):
    """Извлекает весь текст из документа."""
    try:
        doc = Document(docx_path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"[ERROR] Ошибка чтения {docx_path}: {e}")
        return None

def extract_frames_from_indd(indd_path):
    """Открывает InDesign файл и вытаскивает все ID текстовых фреймов и их контент."""
    print(f"[INFO] Извлекаю фреймы из {os.path.basename(indd_path)}...")
    jsx_path = os.path.join(BASE_DIR, "temp_training_extract.jsx")
    json_path = os.path.join(BASE_DIR, "temp_frames.json")
    
    indd_js = indd_path.replace("\\", "/")
    json_js = json_path.replace("\\", "/")
    
    jsx_code = f"""
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
try {{
    var inddFile = new File("{indd_js}");
    var doc = app.open(inddFile);
    
    var jsonStrs = [];
    for (var i = 0; i < doc.textFrames.length; i++) {{
        var rawText = doc.textFrames[i].contents;
        var cleanText = rawText.replace(/\\\\/g, '\\\\\\\\').replace(/"/g, '\\\\"').replace(/\\n/g, '\\\\n').replace(/\\r/g, '\\\\n').replace(/\\t/g, ' ');
        cleanText = cleanText.substring(0, 400); 
        jsonStrs.push('{{"id":"' + i + '", "text":"' + cleanText + '"}}');
    }}
    
    var outFile = new File("{json_js}");
    outFile.encoding = "UTF-8";
    outFile.open("w");
    outFile.write("[" + jsonStrs.join(",") + "]");
    outFile.close();
    
    doc.close(SaveOptions.NO);
}} catch(e) {{
    var errFile = new File("{BASE_DIR}/indesign_error.txt".replace("\\\\", "/"));
    errFile.encoding = "UTF-8";
    errFile.open("w");
    errFile.write(e.toString() + "\\nLine: " + e.line);
    errFile.close();
}} finally {{
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
}}
"""
    with open(jsx_path, 'w', encoding='utf-8') as f:
        f.write(jsx_code)
        
    subprocess.run(['osascript', '-e', f'tell application "Adobe InDesign 2026" to do script POSIX file "{jsx_path}" language javascript'], capture_output=True)
    os.remove(jsx_path)
    
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                print(f"[ERROR] Не удалось прочитать JSON из {indd_path}")
                os.remove(json_path)
                return None
        os.remove(json_path)
        
        data = [item for item in data if item['text'].strip()]
        
        # Превращаем в словарь {"id": "text", ...} для удобства обучения
        frames_dict = {}
        for item in data:
            frames_dict[item["id"]] = item["text"]
        return frames_dict
    else:
        err_log = os.path.join(BASE_DIR, 'indesign_error.txt')
        if os.path.exists(err_log):
            with open(err_log, 'r', encoding='utf-8') as f:
                print(f"[InDesign ERROR] {f.read()}")
        print(f"[ERROR] Не удалось получить структуру фреймов для {indd_path}")
        return None

def main():
    if not os.path.exists(TRAINING_DIR):
        print(f"[ERROR] Папка {TRAINING_DIR} не найдена.")
        return

    dataset = []
    
    folders = [f.path for f in os.scandir(TRAINING_DIR) if f.is_dir()]
    print(f"[INFO] Найдено {len(folders)} папок для обучения.")
    
    for folder in folders:
        folder_name = os.path.basename(folder)
        docx_files = glob.glob(os.path.join(folder, '*.docx'))
        indd_files = glob.glob(os.path.join(folder, '*.indd'))
        
        if not docx_files or not indd_files:
            print(f"[WARNING] Папка {folder_name} содержит неполную пару (DOCX: {len(docx_files)}, INDD: {len(indd_files)}). Пропуск.")
            continue
            
        docx_path = docx_files[0]
        indd_path = indd_files[0]
        
        print(f"\\n--- Анализирую пару из {folder_name} ---")
        
        raw_text = parse_docx_raw(docx_path)
        if not raw_text:
            continue
            
        frames = extract_frames_from_indd(indd_path)
        if not frames:
            continue
            
        dataset.append({
            "source_docx": raw_text,
            "target_frames": frames
        })
        print(f"[SUCCESS] Добавлено в датасет.")
        
    print(f"\\n[INFO] Анализ завершен. Успешно собрано примеров: {len(dataset)}")
    
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(dataset, f, ensure_ascii=False, indent=2)
    print(f"[INFO] Датасет сохранен в {OUTPUT_JSON}")

if __name__ == "__main__":
    main()
