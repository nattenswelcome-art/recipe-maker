import os
import glob
import subprocess
from docx import Document

# Пути к папкам (абсолютные или относительные текущего скрипта)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')

TEMPLATE_NAME = 'recipe_template.indd'
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, TEMPLATE_NAME)

def check_environment():
    """Проверяет наличие необходимых папок и шаблона."""
    for directory in [INPUT_DIR, OUTPUT_DIR, TEMPLATES_DIR]:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"[INFO] Создана папка: {directory}")

    if not os.path.exists(TEMPLATE_PATH):
        print(f"[ERROR] Шаблон не найден: {TEMPLATE_PATH}")
        print("Пожалуйста, поместите InDesign шаблон в папку templates/.")
        return False
    return True

def parse_docx(docx_path):
    """
    Извлекает текст из Word-документа.
    Предполагается, что первый абзац — Заголовок (Title), остальное — Ингредиенты/Текст (Body).
    Также можно усложнить логику, ища определенные стили в Word.
    """
    try:
        doc = Document(docx_path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        
        if not paragraphs:
            return {"title": "Без названия", "body": ""}
            
        title = paragraphs[0]
        body = "\r".join(paragraphs[1:]) # В InDesign перевод каретки это \r
        
        # Экранирование кавычек для передачи в JavaScript (ExtendScript)
        title = title.replace('"', '\\"').replace("'", "\\'")
        body = body.replace('"', '\\"').replace("'", "\\'")
        
        return {"title": title, "body": body}
    except Exception as e:
        print(f"[ERROR] Ошибка чтения {docx_path}: {e}")
        return None

def generate_jsx(recipe_name, text_data, image_path, output_indd, output_pdf):
    """Генерирует JSX-скрипт для InDesign с зашитыми путями и данными."""
    title = text_data.get("title", "")
    body = text_data.get("body", "")
    
    # Экранируем пути для JS (AppleScript иногда ругается на обратные слеши, используем прямые)
    template_js = TEMPLATE_PATH.replace("\\", "/")
    image_js = image_path.replace("\\", "/")
    output_indd_js = output_indd.replace("\\", "/")
    output_pdf_js = output_pdf.replace("\\", "/")

    jsx_code = f"""
// --- Auto-generated InDesign Script ---
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

try {{
    var templateFile = new File("{template_js}");
    if (!templateFile.exists) throw new Error("Шаблон не найден: " + templateFile.fsName);
    
    var doc = app.open(templateFile);
    
    // 1. Вставка текста по Paragraph Styles
    // Мы ищем текстовые фреймы, заглядываем в их первый абзац и проверяем примененный стиль.
    var textFrames = doc.allPageItems;
    for (var i = 0; i < textFrames.length; i++) {{
        var item = textFrames[i];
        if (item instanceof TextFrame && item.parentStory.paragraphs.length > 0) {{
            var pStyleName = item.parentStory.paragraphs[0].appliedParagraphStyle.name;
            
            if (pStyleName === "Title" || pStyleName === "Заголовок") {{
                item.contents = "{title}";
            }} else if (pStyleName === "Ingredients" || pStyleName === "Ингредиенты" || pStyleName === "Body") {{
                item.contents = "{body}";
            }}
        }}
    }}
    
    // 2. Вставка картинки (ищем по Script Label 'Photo' или берем самый большой Rectangle)
    var imageFile = new File("{image_js}");
    if (imageFile.exists) {{
        var photoFrame = null;
        
        // Пытаемся найти по ярлыку (Script Label)
        for (var i = 0; i < doc.allPageItems.length; i++) {{
            if (doc.allPageItems[i].label === "Photo" && doc.allPageItems[i] instanceof Rectangle) {{
                photoFrame = doc.allPageItems[i];
                break;
            }}
        }}
        
        // Если не нашли по ярлыку, берем самый большой Rectangle на документе
        if (!photoFrame) {{
            var maxArea = 0;
            for (var i = 0; i < doc.rectangles.length; i++) {{
                var rect = doc.rectangles[i];
                // bounds: [y1, x1, y2, x2]
                var b = rect.geometricBounds;
                var area = (b[2] - b[0]) * (b[3] - b[1]);
                if (area > maxArea) {{
                    maxArea = area;
                    photoFrame = rect;
                }}
            }}
        }}
        
        if (photoFrame) {{
            photoFrame.place(imageFile);
            photoFrame.fit(FitOptions.FILL_PROPORTIONALLY);
            photoFrame.fit(FitOptions.CENTER_CONTENT);
        }}
    }}
    
    // Проверка на Overset Text
    for (var i = 0; i < doc.textFrames.length; i++) {{
        if (doc.textFrames[i].overflows) {{
            // В идеале можно авто-уменьшить шрифт, но для начала просто выведем warning.
            // Добавим label, чтобы потом считать его из питона, но проще просто сделать fit.
            // doc.textFrames[i].fit(FitOptions.FRAME_TO_CONTENT); // Может сломать верстку
        }}
    }}

    // 3. Сохранение файла
    var outInddFile = new File("{output_indd_js}");
    doc.save(outInddFile);
    
    // 4. Экспорт в PDF
    var outPdfFile = new File("{output_pdf_js}");
    var pdfPreset = app.pdfExportPresets.item("[High Quality Print]"); // Пресет по умолчанию
    if (!pdfPreset.isValid) {{
        pdfPreset = app.pdfExportPresets.firstItem(); // Берем первый если нет HQ Print
    }}
    doc.exportFile(ExportFormat.PDF_TYPE, outPdfFile, false, pdfPreset);
    
    doc.close(SaveOptions.NO);

}} catch (e) {{
    // В случае ошибки записываем ее в текстовый файл, чтобы питон мог ее прочитать
    var errFile = new File("{BASE_DIR}/indesign_error.txt".replace("\\\\", "/"));
    errFile.open("w");
    errFile.write(e.toString() + "\\nLine: " + e.line);
    errFile.close();
}} finally {{
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
}}
"""
    return jsx_code

def run_indesign_script(jsx_path):
    """Запускает JSX скрипт через osascript (AppleScript)."""
    # Удаляем старый лог ошибок, если был
    err_log = os.path.join(BASE_DIR, 'indesign_error.txt')
    if os.path.exists(err_log):
        os.remove(err_log)

    applescript = f'''
    tell application "Adobe InDesign 2024" -- или просто "Adobe InDesign"
        do script POSIX file "{jsx_path}" language javascript
    end tell
    '''
    
    try:
        # Popen позволяет выполнить процесс. osascript -e
        result = subprocess.run(['osascript', '-e', applescript], capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"[ERROR] Ошибка выполнения AppleScript (возможно нет прав или не запущен InDesign): {result.stderr}")
            return False
            
        # Проверяем не оставил ли InDesign лог с ошибкой
        if os.path.exists(err_log):
            with open(err_log, 'r', encoding='utf-8') as f:
                error_msg = f.read()
            print(f"[InDesign ERROR] Скрипт InDesign завершился с ошибкой:\n{error_msg}")
            return False
            
        return True
    except Exception as e:
        print(f"[ERROR] Сбой запуска osascript: {e}")
        return False

def main():
    if not check_environment():
        return

    # Ищем все DOCX в папке input
    docx_files = glob.glob(os.path.join(INPUT_DIR, '*.docx'))
    if not docx_files:
        print("[INFO] В папке input/ нет .docx файлов для обработки.")
        return

    for docx_path in docx_files:
        filename_base = os.path.splitext(os.path.basename(docx_path))[0]
        print(f"\n[INFO] Обрабатываю рецепт: {filename_base}")
        
        # Ищем парную картинку (jpg, jpeg, png)
        image_path = None
        for ext in ['.jpg', '.jpeg', '.png']:
            possible_path = os.path.join(INPUT_DIR, f"{filename_base}{ext}")
            if os.path.exists(possible_path):
                image_path = possible_path
                break
                
        if not image_path:
            print(f"[WARNING] Для {filename_base}.docx не найдено фото с таким же именем. Пропуск.")
            continue

        text_data = parse_docx(docx_path)
        if not text_data:
            continue
            
        output_indd = os.path.join(OUTPUT_DIR, f"{filename_base}.indd")
        output_pdf = os.path.join(OUTPUT_DIR, f"{filename_base}.pdf")
        
        jsx_code = generate_jsx(filename_base, text_data, image_path, output_indd, output_pdf)
        jsx_file_path = os.path.join(BASE_DIR, f"temp_{filename_base}.jsx")
        
        with open(jsx_file_path, 'w', encoding='utf-8') as f:
            f.write(jsx_code)
            
        print(f"[INFO] Запускаю InDesign для генерации {filename_base}...")
        success = run_indesign_script(jsx_file_path)
        
        if success:
            print(f"[SUCCESS] Готово! Сохранено: output/{filename_base}.pdf")
        
        # Удаляем временный js скрипт
        if os.path.exists(jsx_file_path):
            os.remove(jsx_file_path)

if __name__ == "__main__":
    main()
