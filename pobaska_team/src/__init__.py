from puzzle_logger import log_decorator, window_logger
import win32com.client as win32
import os


WORD_EXTENSIONS = {'.doc', '.docx', '.dot', '.dotx', '.docm', '.dotm'}
def is_word_file(file_path):
    return os.path.splitext(file_path)[1].lower() in WORD_EXTENSIONS


@window_logger
@log_decorator
def format_word_document(input_path, output_path=None, font_name="Calibri", font_size=12, replacements=None,
                         puzzle_logger_path=None, **kwargs):
    """
    Основная логика форматирования Word.
    """
    word = None
    doc = None
    try:
        abs_path = os.path.abspath(input_path)
        if not is_word_file(abs_path):
            raise Exception(f"Файл '{input_path}' не является документом Word (неверное расширение)")

        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(input_path))
        for para in doc.Paragraphs:
            para.Range.Font.Reset()
            para.Range.ParagraphFormat.Reset()
        doc.Range().Style = word.ActiveDocument.Styles("Обычный")

        # Применение шрифта
        doc.Range().Font.Name = font_name
        doc.Range().Font.Size = font_size


        # Форматирование таблиц
        for table in doc.Tables:
            table.Borders.Enable = True
            for border in table.Borders:
                border.LineStyle = 1
            table.Borders(-8).LineStyle = 0
            table.Borders(-7).LineStyle = 0
            table.Range.ParagraphFormat.Alignment = 1
            table.Range.Cells.VerticalAlignment = 1
            # Заголовочная строка
            if table.Rows.Count > 0:
                header_row = table.Rows(1)
                header_row.Range.Font.Bold = True
                header_row.Range.Shading.BackgroundPatternColor = 12632256

        # Обработка плейсхолдеров
        if replacements and isinstance(replacements, dict):
            for key, value in replacements.items():
                find_str = f"{{{{{key}}}}}"
                doc.Range().Find.Execute(FindText=find_str, ReplaceWith=str(value), Replace=2)

        save_path = output_path if output_path else input_path
        doc.SaveAs(os.path.abspath(save_path))
        return "Успешно отформатировано"

    except Exception as e:
        raise Exception(f"Ошибка при работе с Word: {str(e)}")
    finally:
        if doc:
            doc.Close()
        if word:
            word.Quit()
