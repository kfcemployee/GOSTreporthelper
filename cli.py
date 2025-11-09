import sys
import os
import argparse
from utils.parsers import DocParser, MdParser
from utils.converters import Converter, MarkdownConverter
from docx import Document


def process_file(input_file, output_file=None, force=False):
    try:
        if not output_file:
            output_file = "output.docx"

        if os.path.exists(output_file) and not force:
            print(f"файл {output_file} уже существует. используйте -f для перезаписи.")
            return False

        if input_file.endswith(".txt"):
            with open(input_file, encoding="utf-8") as text:
                data = MdParser(text.read()).parse_()
                d = MarkdownConverter(data, output_file).convert_to_doc()
                p1 = DocParser(d)
                Converter(d, p1.parse(), output_file).start()
        else:
            doc = Document(input_file)
            p1 = DocParser(doc)
            Converter(doc, p1.parse(), output_file).start()

        print(f"готово: {output_file}")
        return True

    except Exception as e:
        print(f"ошибка: {e}")
        return False


def main():
    if len(sys.argv) == 2 and os.path.isfile(sys.argv[1]):
        input_file = sys.argv[1]
        if input_file.lower().endswith(('.docx', '.txt')):
            print(f"обработка: {os.path.basename(input_file)}")
            success = process_file(input_file)
            if success:
                input("нажмите Enter для выхода...")
            else:
                input("нажмите Enter для выхода...")
        else:
            print("неподдерживаемый файл")
            input("нажмите Enter для выхода...")
        return

    pr = argparse.ArgumentParser(
        description='GOST Report Helper - форматирование документов по ГОСТу',
        add_help=False
    )

    pr.add_argument(
        'input_file',
        nargs='?',
        help='путь к входному файлу (.docx или .txt)'
    )

    pr.add_argument(
        '-o', '--output',
        help='путь к выходному файлу (по умолчанию: output.docx)',
        default=None
    )

    pr.add_argument(
        '-h', '--help',
        action='store_true',
        help='показать справку'
    )

    args = pr.parse_args()

    if args.help or not args.input_file:
        print("GOST report helper - форматирование документов по ГОСТу\n")
        print("Использование:")
        print("Перетащите файл .docx или .txt на GOSTreporthelper.exe")
        print("Или используйте командную строку:\n")
        print("GOSTFormatter.exe файл.docx")
        print("GOSTFormatter.exe файл.txt -o результат.docx")
        print("GOSTFormatter.exe файл.docx -f")
        print("\nОпции:")
        print("-o, --output ФАЙЛ   Выходной файл")
        print("-h, --help          Показать эту справку")
        input("Нажмите Enter для выхода... ")
        return

    if not os.path.exists(args.input_file):
        print(f"файл не найден: {args.input_file}")
        return

    if not args.input_file.lower().endswith(('.docx', '.txt')):
        print("неподдерживаемый файл")
        return

    process_file(args.input_file, args.output, args.force)

main()