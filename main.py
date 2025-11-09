from parsers import DocParser, MdParser
from converters import Converter, MarkdownConverter

file_path = r""

if not file_path:
    file_path = input("Введите путь к файлу: ")

o_path = "__output.docx"

if file_path.endswith(".txt"):
    with open(file_path, encoding="utf-8") as text:
        data = MdParser(text.read()).parse_()
        d = MarkdownConverter(data, o_path).convert_to_doc()
        p1 = DocParser(d)
        Converter(d, p1.parse(), o_path).start()
elif file_path.endswith(".docx"):
    from docx import Document
    try:
        doc = Document(file_path)
        p1 = DocParser(doc)
        Converter(doc, p1.parse(), o_path).start()
    except Exception as e:
        print("Проверьте, точно ли ваш файл является docx. ", e)

else:
    print("Неподдерживаемый тип файла: может быть только docx или txt.")
