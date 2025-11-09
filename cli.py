import argparse
import sys
import os
from utils.parsers import DocParser, MdParser
from utils.converters import Converter, MarkdownConverter
from docx import Document

def main():
    if len(sys.argv) == 1 or '--help' in sys.argv or '-h' in sys.argv:
        print("Использование: python cli.py <файл> [опции]\n")
        print("Обязательные аргументы:")
        print("[файл]                Входной файл (.docx или .txt)\n")
        print("Опции:")
        print("-o, --output ФАЙЛ     Выходной файл (по умолчанию: output.docx)")
        print("-f, --force           Перезаписать существующий файл")
        print("-h, --help            Показать эту справку\n")
        print("Примеры:")
        print("python cli.py документ.docx")
        print("python cli.py текст.txt -o отчет.docx -v")
        print("python cli.py файл.docx -f")
        sys.exit(0)

    pr = argparse.ArgumentParser(
        description='GOST Report Helper - форматирование документов по ГОСТу',
        add_help=False
    )

    pr.add_argument(
        'input_file',
        help='Путь к входному файлу (.docx или .txt)'
    )

    pr.add_argument(
        '-o', '--output',
        help='Путь к выходному файлу (по умолчанию: output.docx)',
        default='output.docx'
    )

    pr.add_argument(
        '-f', '--force',
        help='Перезаписать выходной файл если существует',
        action='store_true'
    )

    pr.add_argument(
        '-h', '--help',
        help='Показать справку',
        action='store_true'
    )

    args = pr.parse_args()

    if not os.path.exists(args.input_file):
        print(f"ошибка: файл '{args.input_file}' не найден")
        sys.exit(1)

    if not args.input_file.lower().endswith(('.docx', '.txt')):
        print("ошибка: неподдерживаемый формат файла")
        sys.exit(1)

    if os.path.exists(args.output) and not args.force:
        print(f"ошибка: файл '{args.output}' уже существует,")
        print("используйте -f для перезаписи")
        sys.exit(1)

    try:
        if args.input_file.endswith(".txt"):
            with open(args.input_file, encoding="utf-8") as text:
                data = MdParser(text.read()).parse_()
                d = MarkdownConverter(data, args.output).convert_to_doc()
                p1 = DocParser(d)
                Converter(d, p1.parse(), args.output).start()
        else:
            doc = Document(args.input_file)
            p1 = DocParser(doc)
            Converter(doc, p1.parse(), args.output).start()

        print(f"Готово: {args.output}")

    except Exception as e:
        print(f"ошибка при обработке: {e}")
        sys.exit(1)

main()
