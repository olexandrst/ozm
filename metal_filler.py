"""
Заповнення колонок "Наявність металу", "Назва матеріалу" та
"Вміст металу в одиниці виміру ОЗМ, відмінної від вагової"
у файлі .xlsb (Excel Binary) через Azure OpenAI.

Запуск (Windows, потрібен встановлений Excel):
    python metal_filler.py 1.xlsb
    python metal_filler.py 2.xlsb --sheet 1 --batch 20 --save-every 1000

Залежності:
    pip install pywin32 openai tenacity python-dotenv

Змінні середовища (.env або system):
    AZURE_OPENAI_ENDPOINT      = https://<your-resource>.openai.azure.com
    AZURE_OPENAI_API_KEY       = <key>
    AZURE_OPENAI_API_VERSION   = 2024-10-21        (або новіша)
    AZURE_OPENAI_DEPLOYMENT    = gpt-5.2           (назва deployment-а)
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from dataclasses import dataclass
from typing import Optional

import pythoncom
import win32com.client as win32
from dotenv import load_dotenv
from openai import AzureOpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from openai import APIError, RateLimitError, APITimeoutError


# ---------- Промпт ----------------------------------------------------------

SYSTEM_PROMPT = """\
Ти — інженер-матеріалознавець на металургійному підприємстві. Аналізуєш
найменування ОЗМ (основних засобів / матеріалів) українською та російською
мовами і визначаєш вміст металу.

Для кожної позиції повертай JSON-обʼєкт з полями:

1. "has_metal" (bool) — чи містить ОЗМ метал як суттєвий компонент.
   - true: металопрокат, кріплення, інструмент, кабель, обладнання з металу,
     металобрухт, дріт, труби, листи, балки, чавунні/сталеві деталі тощо.
   - false: деревина, тканина, папір, пластик, ГСМ, гума, хімія, бетон,
     ЗІЗ без металу, продукти харчування тощо.

2. "material" (string|null) — НАЗВА основного металу (того, що має найбільшу
   масову частку). Допустимі значення (у називному відмінку, українською):
     "сталь", "чавун", "мідь", "алюміній", "цинк", "олово", "свинець",
     "латунь", "бронза", "нікель", "титан", "срібло", "золото", "метал".
   Якщо точно не визначити — "метал". Якщо has_metal=false — null.

3. "metal_kg_per_unit" (number|null) — оціночна МАСА металу В КІЛОГРАМАХ
   в ОДНІЙ одиниці виміру ОЗМ (одиниця подається у полі "unit":
   ШТ=штука, М=метр, М3=кубометр, УПК=упаковка, КМП=комплект, КГ=кілограм,
   Т=тонна).
   Правила:
     - Для Т: 1 тонна = 1000 кг чистого металу, якщо ОЗМ — це сам метал
       (брухт, прокат). Інакше — частка від 1000 кг.
     - Для КГ: 1 кг ОЗМ = (масова частка металу) * 1. Якщо ОЗМ — сам метал,
       значення = 1.0.
     - Для ШТ/М/М3/УПК/КМП: оціни типову вагу металу в одиниці виходячи з
       назви, типорозміру (якщо вказаний у назві) і галузевого досвіду.
     - Якщо has_metal=false — null.
     - Якщо has_metal=true, але оцінити неможливо — постав консервативну
       оцінку (>0), не null.

Відповідай ВИКЛЮЧНО валідним JSON у форматі:
{"results":[{"i":<index>, "has_metal":..., "material":..., "metal_kg_per_unit":...}, ...]}

Не додавай пояснень, markdown чи коментарів. Порядок елементів — як на вході.
"""

USER_TEMPLATE = (
    "Проаналізуй позиції. Поверни results у тому ж порядку:\n{items_json}"
)


# ---------- Колонки в таблиці ----------------------------------------------

COL_INDEX = 1        # A - Індекс
COL_NAME = 2         # B - Назва
COL_UNIT = 3         # C - Одиниці
COL_HAS_METAL = 4    # D - Наявність металу
COL_MATERIAL = 5     # E - Назва матеріалу
COL_METAL_QTY = 6    # F - Вміст металу


# ---------- LLM -------------------------------------------------------------

@dataclass
class Row:
    excel_row: int
    name: str
    unit: str


def build_client() -> tuple[AzureOpenAI, str]:
    load_dotenv()
    endpoint = os.environ["AZURE_OPENAI_ENDPOINT"]
    api_key = os.environ["AZURE_OPENAI_API_KEY"]
    api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-10-21")
    deployment = os.environ["AZURE_OPENAI_DEPLOYMENT"]
    client = AzureOpenAI(
        azure_endpoint=endpoint, api_key=api_key, api_version=api_version
    )
    return client, deployment


@retry(
    reraise=True,
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=2, max=60),
    retry=retry_if_exception_type((RateLimitError, APITimeoutError, APIError)),
)
def call_llm(client: AzureOpenAI, deployment: str, batch: list[Row]) -> list[dict]:
    items = [{"i": i, "name": r.name, "unit": r.unit} for i, r in enumerate(batch)]
    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": USER_TEMPLATE.format(items_json=json.dumps(items, ensure_ascii=False))},
        ],
        response_format={"type": "json_object"},
        temperature=0,
    )
    payload = json.loads(resp.choices[0].message.content)
    results = payload.get("results", [])
    # Вирівнюємо за індексом i, на випадок порушеного порядку
    by_i = {int(r["i"]): r for r in results}
    out = []
    for i in range(len(batch)):
        r = by_i.get(i, {"has_metal": False, "material": None, "metal_kg_per_unit": None})
        out.append(r)
    return out


# ---------- Excel (win32com, читання/запис .xlsb) --------------------------

class ExcelBook:
    def __init__(self, path: str, sheet: str):
        self.path = os.path.abspath(path)
        self.sheet_name = sheet
        self.app = None
        self.wb = None
        self.ws = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = win32.gencache.EnsureDispatch("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.ScreenUpdating = False
        self.wb = self.app.Workbooks.Open(self.path)
        self.ws = self.wb.Worksheets(self.sheet_name)
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=False)
        finally:
            if self.app is not None:
                self.app.Quit()
            pythoncom.CoUninitialize()

    def last_row(self) -> int:
        # xlUp = -4162
        return int(self.ws.Cells(self.ws.Rows.Count, COL_NAME).End(-4162).Row)

    def read_block(self, start_row: int, end_row: int) -> list[tuple[str, str, object]]:
        """Повертає список (name, unit, has_metal_value) для рядків [start..end]."""
        rng = self.ws.Range(
            self.ws.Cells(start_row, COL_NAME),
            self.ws.Cells(end_row, COL_HAS_METAL),
        ).Value
        if end_row == start_row:
            rng = (rng,)
        out = []
        for tup in rng:
            name, unit, has_metal = tup[0], tup[1], tup[2]
            out.append((name, unit, has_metal))
        return out

    def write_results(self, rows: list[Row], results: list[dict]):
        """Запис блоком D:F для діапазону рядків."""
        if not rows:
            return
        # Сортуємо за excel_row на випадок
        pairs = sorted(zip(rows, results), key=lambda p: p[0].excel_row)
        # Пишемо по неперервних сегментах
        seg_start = 0
        for k in range(1, len(pairs) + 1):
            if k == len(pairs) or pairs[k][0].excel_row != pairs[k - 1][0].excel_row + 1:
                segment = pairs[seg_start:k]
                first_row = segment[0][0].excel_row
                last_row = segment[-1][0].excel_row
                values = []
                for _, r in segment:
                    has = bool(r.get("has_metal"))
                    mat = r.get("material") if has else None
                    qty = r.get("metal_kg_per_unit") if has else None
                    values.append([
                        "так" if has else "ні",
                        mat if mat is not None else "",
                        qty if qty is not None else "",
                    ])
                self.ws.Range(
                    self.ws.Cells(first_row, COL_HAS_METAL),
                    self.ws.Cells(last_row, COL_METAL_QTY),
                ).Value = values
                seg_start = k

    def save(self):
        # 50 = xlExcel12 (.xlsb)
        self.wb.Save()


# ---------- Основний цикл ---------------------------------------------------

def is_empty(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def find_first_empty(book: ExcelBook, last_row: int, header_rows: int) -> int:
    """Знаходимо перший рядок з пустою "Наявність металу"."""
    # Читаємо колонку D одним блоком — швидко
    rng = book.ws.Range(
        book.ws.Cells(header_rows + 1, COL_HAS_METAL),
        book.ws.Cells(last_row, COL_HAS_METAL),
    ).Value
    if last_row == header_rows + 1:
        rng = (rng,) if not isinstance(rng, tuple) else rng
        rng = [(rng[0],)] if not isinstance(rng[0], tuple) else rng
    for i, tup in enumerate(rng):
        v = tup[0] if isinstance(tup, tuple) else tup
        if is_empty(v):
            return header_rows + 1 + i
    return last_row + 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file", help="Шлях до .xlsb файлу")
    ap.add_argument("--sheet", default=None, help="Назва аркуша (за замовч. — перший)")
    ap.add_argument("--header-rows", type=int, default=1, help="К-сть рядків заголовку")
    ap.add_argument("--batch", type=int, default=20, help="Позицій у одному запиті до LLM")
    ap.add_argument("--save-every", type=int, default=1000, help="Зберігати кожні N рядків")
    ap.add_argument("--delay", type=float, default=0.4, help="Пауза між запитами, сек")
    ap.add_argument("--limit", type=int, default=0, help="Обробити максимум N рядків (0 = всі)")
    args = ap.parse_args()

    client, deployment = build_client()

    sheet_name = args.sheet
    if sheet_name is None:
        # Відкриваємо щоб дізнатись назву першого аркуша
        pythoncom.CoInitialize()
        app = win32.gencache.EnsureDispatch("Excel.Application")
        app.Visible = False
        wb = app.Workbooks.Open(os.path.abspath(args.file))
        sheet_name = wb.Worksheets(1).Name
        wb.Close(SaveChanges=False)
        app.Quit()
        pythoncom.CoUninitialize()

    with ExcelBook(args.file, sheet_name) as book:
        last = book.last_row()
        start = find_first_empty(book, last, args.header_rows)
        if start > last:
            print("Всі рядки вже заповнені.")
            return
        print(f"Старт з рядка {start}, останній рядок {last} (всього {last - start + 1}).")

        processed_since_save = 0
        total_processed = 0
        cur = start

        while cur <= last:
            block_end = min(cur + args.batch - 1, last)
            raw = book.read_block(cur, block_end)
            rows: list[Row] = []
            for offset, (name, unit, has_metal) in enumerate(raw):
                excel_row = cur + offset
                if not is_empty(has_metal):
                    continue  # вже заповнено — пропускаємо
                if is_empty(name):
                    continue
                rows.append(Row(excel_row=excel_row, name=str(name), unit=str(unit or "")))

            if rows:
                try:
                    results = call_llm(client, deployment, rows)
                    book.write_results(rows, results)
                except Exception as e:
                    print(f"[!] Помилка на рядках {rows[0].excel_row}..{rows[-1].excel_row}: {e}",
                          file=sys.stderr)
                processed_since_save += len(rows)
                total_processed += len(rows)
                print(f"  оброблено {total_processed} (рядки {rows[0].excel_row}..{rows[-1].excel_row})")
                time.sleep(args.delay)

            if processed_since_save >= args.save_every:
                print(f"  -> збереження ({processed_since_save} нових)")
                book.save()
                processed_since_save = 0

            if args.limit and total_processed >= args.limit:
                break

            cur = block_end + 1

        if processed_since_save > 0:
            print(f"  -> фінальне збереження ({processed_since_save})")
            book.save()
        print(f"Готово. Оброблено {total_processed} рядків.")


if __name__ == "__main__":
    main()
