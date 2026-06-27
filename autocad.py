# Скрипт читает данные из Excel и подставляет их в шаблоны AutoCAD и Word

import win32com.client
import pandas as pd
import os
import time
import sys


# ========== НАСТРОЙКИ ==========
# Получаем путь к папке, где находится скрипт
if getattr(sys, 'frozen', False):
    SCRIPT_FOLDER = os.path.dirname(sys.executable)
else:
    SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))

# Указываем имена файлов (они лежат в той же папке)
EXCEL_FILE = os.path.join(SCRIPT_FOLDER, "шаблон.xlsx")
TEMPLATE_DWG = os.path.join(SCRIPT_FOLDER, "шаблон.dwg")
TEMPLATE_DOCX = os.path.join(SCRIPT_FOLDER, "шаблон.docx")
OUTPUT_FOLDER_DWG = os.path.join(SCRIPT_FOLDER, "чертежи")
OUTPUT_FOLDER_DOCX = os.path.join(SCRIPT_FOLDER, "документы")
SHEET_NAME = "Лист1" # название листа в Excel

# ========== СОЗДАНИЕ ПАПОК ДЛЯ РЕЗУЛЬТАТОВ ==========
if not os.path.exists(OUTPUT_FOLDER_DWG):
    os.makedirs(OUTPUT_FOLDER_DWG)

if not os.path.exists(OUTPUT_FOLDER_DOCX):
    os.makedirs(OUTPUT_FOLDER_DOCX)

# ========== ПОДКЛЮЧЕНИЕ К AUTOCAD ==========
try: 
    # Пытаемся подключиться к уже открытому AutoCAD
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print("\nПодключено к запущенному AutoCAD")
except:
    # Если AutoCAD не запущен - запускаем его
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    print("\nAutoCAD запущен")

# ========== ПОДКЛЮЧЕНИЕ К WORD ==========
try:
    # Пытаемся подключиться к уже открытому Word
    word = win32com.client.GetActiveObject("Word.Application")
    print("Подключено к запущенному Word")
except:
    # Если Word не запущен - запускаем его
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # False - работать в фоне, True - показывать окно
    print("Word запущен")

# ========== ЧТЕНИЕ EXCEL ==========
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME) # Читаем Excel-файл в таблицу

# ========== ФУНКЦИЯ ДЛЯ ЗАМЕНЫ ТЕКСТА В ЧЕРТЕЖЕ AUTOCAD ==========
def replace_text_in_drawing(doc, replacements):
    """
    Заменяет текст во всех текстовых объектах чертежа AutoCAD.
    
    Параметры:
        doc - открытый документ AutoCAD
        replacements - словарь вида {"маркер": "значение"}
    
    Возвращает:
        количество выполненных замен
    """
    replaced_count = 0 # Счетчик замен
    
    try:
        # Получаем количество объектов в пространстве модели
        count = doc.ModelSpace.Count
    except:
        print("Не удалось получить количество объектов в AutoCAD")
        return 0
    
    for i in range(count):
        # Перебираем все объекты в чертеже по индексу
        try:
            obj = doc.ModelSpace.Item(i) # Получаем объект по индексу
            obj_name = obj.ObjectName # Узнаем тип объекта
            
            if obj_name in ["AcDbText", "AcDbMText"]:
                # Если объект - текст (обычный или многострочный)
                old_text = obj.TextString # Исходный текст
                new_text = old_text # Копия для изменений
                
                for placeholder, value in replacements.items():
                    # проходим по всем маркерам и заменяем их
                    if placeholder in new_text:
                        new_text = new_text.replace(placeholder, str(value))
                
                if new_text != old_text:
                    # Если текст изменился - применяем замену
                    obj.TextString = new_text # Устанавливаем новый текст
                    replaced_count += 1
                    
        except Exception as e:
            # Пропускаем проблемные объекты
            continue
    
    return replaced_count

# ========== ФУНКЦИЯ ДЛЯ ЗАМЕНЫ ТЕКСТА В ДОКУМЕНТЕ WORD ==========
def replace_text_in_word(doc, replacements):
    """
    Заменяет текст во всем документе Word.
    
    Параметры:
        doc - открытый документ Word
        replacements - словарь вида {"маркер": "значение"}
    
    Возвращает:
        количество выполненных замен
    """
    replaced_count = 0
    
    for placeholder, value in replacements.items():
        for story in doc.StoryRanges:
            story_range = story
            while story_range:
                try:
                    text = story_range.Text
                    if placeholder in text:
                        new_text = text.replace(placeholder, str(value))
                        story_range.Text = new_text
                        replaced_count += 1
                except:
                    pass
                story_range = story_range.NextStoryRange
    
    return replaced_count

# ========== ОБРАБОТКА КАЖДОЙ СТРОКИ ==========
for index, row in df.iterrows():
    print(f"\n{'='*50}")
    print(f"Обработка строки {index + 1} из {len(df)}")
    print('='*50)
    
    # Получение данных из Excel
    nazvanie = str(row['Название'])
    
    if 'МощностьТекст' in df.columns and pd.notna(row['МощностьТекст']):
        moshnost = str(row['МощностьТекст'])
    else:
        moshnost = str(row['Мощность'])
    
    if 'КПДТекст' in df.columns and pd.notna(row['КПДТекст']):
        kpd = str(row['КПДТекст'])
    else:
        kpd = str(row['КПД'])
    
    if 'ДавлениеТекст' in df.columns and pd.notna(row['ДавлениеТекст']):
        davlenie = str(row['ДавлениеТекст'])
    else:
        davlenie = str(row['Давление'])
    
    tip_kotla = str(row['ТипКотла'])
    
    # Создаем безопасное имя файла
    safe_name = nazvanie.replace('"', '').replace('/', '').replace('\\', '').replace(':', '').replace('*', '')
    
    # Подготавливаем словари замен
    replacements_acad = {
        r"\{Название\}": nazvanie,
        r"\{Мощность\}": moshnost,
        r"\{КПД\}": kpd,
        r"\{Давление\}": davlenie,
        r"\{ТипКотла\}": tip_kotla
    }

    replacements_word = {
        "\{Название\}": nazvanie,
        "\{Мощность\}": moshnost,
        "\{КПД\}": kpd,
        "\{Давление\}": davlenie,
        "\{ТипКотла\}": tip_kotla
    }
    
    # ========== ОБРАБОТКА AUTOCAD ==========
    print("\n[AUTOCAD]")
    if os.path.exists(TEMPLATE_DWG):
        doc_acad = None
        try:
            # Открываем шаблон AutoCAD
            doc_acad = acad.Documents.Open(TEMPLATE_DWG)
            print("✓ Шаблон DWG открыт")
            
            # Небольшая пауза для загрузки чертежа
            time.sleep(0.5)
            
            # Заменяем текст в чертеже
            replaced_count = replace_text_in_drawing(doc_acad, replacements_acad)
            
            if replaced_count > 0:
                print(f"✓ Выполнено замен в DWG: {replaced_count}")
            else:
                print("⚠ ВНИМАНИЕ: Замен в DWG не произошло!")
                print("  Проверьте, что в шаблоне есть маркеры: {Название}, {Мощность}, {КПД}, {Давление}, {ТипКотла}")
            
            # Сохраняем чертеж
            output_file_dwg = os.path.join(OUTPUT_FOLDER_DWG, f"{safe_name}.dwg")
            doc_acad.SaveAs(output_file_dwg)
            print(f"✓ DWG сохранен: {os.path.basename(output_file_dwg)}")
            
            # Закрываем чертеж
            doc_acad.Close(False)
            
        except Exception as e:
            print(f"✗ ОШИБКА в AutoCAD: {e}")
            if doc_acad:
                try:
                    doc_acad.Close(False)
                except:
                    pass
    else:
        print(f"✗ Шаблон DWG не найден: {TEMPLATE_DWG}")
    
    # ========== ОБРАБОТКА WORD ==========
    print("\n[WORD]")
    if os.path.exists(TEMPLATE_DOCX):
        doc_word = None
        try:
            # Открываем шаблон Word
            doc_word = word.Documents.Open(TEMPLATE_DOCX)
            print("✓ Шаблон DOCX открыт")
            
            # Небольшая пауза
            time.sleep(0.3)
            
            # Заменяем текст в документе
            replaced_count = replace_text_in_word(doc_word, replacements_word)
            
            if replaced_count > 0:
                print(f"✓ Выполнено замен в DOCX: {replaced_count}")
            else:
                print("⚠ ВНИМАНИЕ: Замен в DOCX не произошло!")
                print("  Проверьте маркеры в шаблоне Word")
            
            # Сохраняем документ
            output_file_docx = os.path.join(OUTPUT_FOLDER_DOCX, f"{safe_name}.docx")
            doc_word.SaveAs(output_file_docx)
            print(f"✓ DOCX сохранен: {os.path.basename(output_file_docx)}")
            
            # Закрываем документ
            doc_word.Close(False)
            
        except Exception as e:
            print(f"✗ ОШИБКА в Word: {e}")
            if doc_word:
                try:
                    doc_word.Close(False)
                except:
                    pass
    else:
        print(f"✗ Шаблон DOCX не найден: {TEMPLATE_DOCX}")
    
    # Небольшая пауза между файлами
    time.sleep(0.5)

# ========== ЗАВЕРШЕНИЕ РАБОТЫ ==========
print("\n" + "="*50)
print("ГОТОВО! Все файлы созданы.")
print(f"Чертежи AutoCAD сохранены в: {OUTPUT_FOLDER_DWG}")
print(f"Документы Word сохранены в: {OUTPUT_FOLDER_DOCX}")

# Закрываем Word (если он был запущен скриптом)
try:
    word.Quit()
    print("Word закрыт")
except:
    pass

input("\nНажмите Enter для выхода...")