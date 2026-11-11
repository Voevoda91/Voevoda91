#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Экспорт комплекта ОПР корпуса в .docx.

Для каждого корпуса: 8 отдельных документов + сводный отчёт (все 8 в одном
файле). Ориентация альбомная, A4, Times New Roman 10 pt — широкие таблицы
карты рисков читаются без обрезки.

Запуск:  .venv/bin/python tools/opr-gen/docx.py
Зависимости: pip install pypandoc-binary python-docx
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BASE = os.path.join(ROOT, 'docs', 'opr-mbdou-ds5')
TOOL = os.path.join(ROOT, 'tools', 'md2docx.py')

# (имя файла .md, имя файла .docx)
DOCS = [
    ('00-opis-komplekta-i-metodika.md', '00_Опись_комплекта_и_методика.docx'),
    ('01-prikazy.md', '01_Приказы_по_ОПР.docx'),
    ('02-polozhenie-upravlenie-profriskami.md',
     '02_Положение_об_управлении_профрисками.docx'),
    ('03-perechen-rabochih-mest.md', '03_Перечень_рабочих_мест.docx'),
    ('04-karta-opasnostey-i-profriskov.md', '04_Карта_опасностей_и_рисков.docx'),
    ('05-plan-meropriyatiy.md', '05_План_мероприятий.docx'),
    ('06-svodnaya-vedomost-i-pokazateli.md',
     '06_Сводная_ведомость_и_показатели.docx'),
    ('07-oznakomlenie-i-monitoring.md', '07_Ознакомление_и_мониторинг.docx'),
]

BUILDINGS = [
    ('frunze', 'Отчёт_ОПР_корпус_1_Фрунзе_87.docx'),
    ('gagarina', 'Отчёт_ОПР_корпус_2_Гагарина_9А.docx'),
    ('zadorozhnaya', 'Отчёт_ОПР_корпус_3_Задорожная_14.docx'),
]


def convert(src, dst, venv_python):
    subprocess.run([venv_python, TOOL, src, dst, 'land', '10'], check=True)


def main():
    venv_python = sys.executable
    ok = fail = 0
    for bld, summary_name in BUILDINGS:
        src_dir = os.path.join(BASE, bld)
        out_dir = os.path.join(src_dir, 'word')
        os.makedirs(out_dir, exist_ok=True)
        parts = []
        for md, docx in DOCS:
            src = os.path.join(src_dir, md)
            dst = os.path.join(out_dir, docx)
            try:
                convert(src, dst, venv_python)
                ok += 1
            except subprocess.CalledProcessError as e:
                print('  ОШИБКА %s: %s' % (src, e))
                fail += 1
                continue
            with open(src, encoding='utf-8') as f:
                parts.append(f.read().rstrip() + '\n')
        # сводный отчёт по корпусу — все 8 документов в одном файле
        combined = '\n\n---\n\n'.join(parts)
        with tempfile.NamedTemporaryFile('w', suffix='.md', delete=False,
                                         encoding='utf-8') as tmp:
            tmp.write(combined)
            tmp_path = tmp.name
        try:
            convert(tmp_path, os.path.join(out_dir, summary_name), venv_python)
            ok += 1
        except subprocess.CalledProcessError as e:
            print('  ОШИБКА сводный отчёт %s: %s' % (bld, e))
            fail += 1
        finally:
            os.unlink(tmp_path)
        print('%s: готово' % bld)
    print('\nУспешно: %d, с ошибками: %d' % (ok, fail))
    return 1 if fail else 0


if __name__ == '__main__':
    sys.exit(main())
