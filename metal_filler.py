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
from openai import APIError, RateLimitError, APITimeoutError, APIConnectionError


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


# ---------- Заголовки колонок (для автодетекту) ----------------------------

# Ключове слово в заголовку -> внутрішнє ім'я.
# Шукаємо по підрядку (без врахування регістру), пробілам, переносам рядка.
HEADER_PATTERNS = {
    "name":       ["назва"],            # ОЗМ
    "unit":       ["одиниц"],           # Одиниці / единицы
    "has_metal":  ["наявність метал", "наявн"],
    "material":   ["назва матеріал", "матеріал"],
    "metal_qty":  ["вміст метал"],
}


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
    stop=stop_after_attempt(15),
    wait=wait_exponential(multiplier=2, min=5, max=180),
    retry=retry_if_exception_type(
        (RateLimitError, APITimeoutError, APIConnectionError, APIError)
    ),
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
    def __init__(self, path: str, sheet):
        self.path = os.path.abspath(path)
        # sheet: None  -> перший аркуш; int -> за індексом (1-based);
        #        str   -> за назвою. Числовий рядок ("1") теж сприймається як назва.
        self.sheet = sheet
        self.app = None
        self.wb = None
        self.ws = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = win32.gencache.EnsureDispatch("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.ScreenUpdating = False
        self.wb = self.app.Workbooks.Open(
            self.path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True
        )
        self.ws = self._resolve_sheet()
        print(f"Аркуш: '{self.ws.Name}' (індекс {self.ws.Index})")
        self.cols = self._detect_columns()
        return self

    def _detect_columns(self) -> dict:
        """Зчитує перший рядок і знаходить індекси колонок за ключовими словами."""
        last_col = int(self.ws.Cells(1, self.ws.Columns.Count).End(-4159).Column)  # xlToLeft
        # Іноді End(xlToLeft) повертає 1 на пустому рядку — підстраховка:
        if last_col < 5:
            last_col = max(last_col, 32)
        header_row = self.ws.Range(
            self.ws.Cells(1, 1), self.ws.Cells(1, last_col)
        ).Value
        if not isinstance(header_row, tuple):
            header_row = (header_row,)
        headers = header_row[0] if isinstance(header_row[0], tuple) else header_row
        cols = {}
        for col_idx, raw in enumerate(headers, start=1):
            if raw is None:
                continue
            norm = " ".join(str(raw).lower().split())
            for key, patterns in HEADER_PATTERNS.items():
                if key in cols:
                    continue
                if any(p in norm for p in patterns):
                    cols[key] = col_idx
                    break
        missing = [k for k in HEADER_PATTERNS if k not in cols]
        if missing:
            shown = {i + 1: headers[i] for i in range(len(headers)) if headers[i]}
            raise RuntimeError(
                f"Не знайдено колонки за заголовками: {missing}. "
                f"Заголовки у файлі: {shown}"
            )
        print(
            f"Колонки: Назва=col{cols['name']}, Одиниці=col{cols['unit']}, "
            f"Наявність=col{cols['has_metal']}, Матеріал=col{cols['material']}, "
            f"Вміст=col{cols['metal_qty']}"
        )
        return cols

    def _resolve_sheet(self):
        sheets = self.wb.Worksheets
        if self.sheet is None:
            return sheets(1)
        if isinstance(self.sheet, int):
            return sheets(self.sheet)
        # str — спершу пробуємо за назвою, потім fallback на індекс,
        # якщо рядок є числом ("1") і аркуша з такою назвою нема.
        try:
            return sheets(self.sheet)
        except pythoncom.com_error:
            if self.sheet.isdigit():
                return sheets(int(self.sheet))
            available = [sheets(i + 1).Name for i in range(sheets.Count)]
            raise RuntimeError(
                f"Аркуш '{self.sheet}' не знайдено. Доступні: {available}"
            )

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=False)
        finally:
            if self.app is not None:
                self.app.Quit()
            pythoncom.CoUninitialize()

    def last_row(self) -> int:
        # xlUp = -4162; визначаємо по колонці "Назва"
        return int(self.ws.Cells(self.ws.Rows.Count, self.cols["name"]).End(-4162).Row)

    def read_block(self, start_row: int, end_row: int) -> list[tuple[str, str, object]]:
        """Повертає (name, unit, has_metal_value) для рядків [start..end]."""
        c_name = self.cols["name"]
        c_unit = self.cols["unit"]
        c_has = self.cols["has_metal"]

        def read_col(col):
            rng = self.ws.Range(
                self.ws.Cells(start_row, col), self.ws.Cells(end_row, col)
            ).Value
            # Excel повертає кортеж кортежів для діапазону, або скалярне значення для 1 комірки
            if end_row == start_row:
                return [rng]
            return [row[0] if isinstance(row, tuple) else row for row in rng]

        names = read_col(c_name)
        units = read_col(c_unit)
        hases = read_col(c_has)
        return list(zip(names, units, hases))

    def write_results(self, rows: list[Row], results: list[dict]):
        """Запис у три колонки (можуть бути несуміжними) — комірка за коміркою."""
        if not rows:
            return
        c_has = self.cols["has_metal"]
        c_mat = self.cols["material"]
        c_qty = self.cols["metal_qty"]
        for row_obj, r in zip(rows, results):
            has = bool(r.get("has_metal"))
            mat = r.get("material") if has else None
            qty = r.get("metal_kg_per_unit") if has else None
            self.ws.Cells(row_obj.excel_row, c_has).Value = "так" if has else "ні"
            self.ws.Cells(row_obj.excel_row, c_mat).Value = mat if mat is not None else ""
            self.ws.Cells(row_obj.excel_row, c_qty).Value = qty if qty is not None else ""

    def save(self):
        # 50 = xlExcel12 (.xlsb)
        self.wb.Save()


# ---------- Основний цикл ---------------------------------------------------

def is_empty(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def find_first_empty(book: ExcelBook, last_row: int, header_rows: int) -> int:
    """Знаходимо перший рядок з пустою колонкою "Наявність металу"."""
    col = book.cols["has_metal"]
    rng = book.ws.Range(
        book.ws.Cells(header_rows + 1, col),
        book.ws.Cells(last_row, col),
    ).Value
    if last_row == header_rows + 1:
        rng = (rng,) if not isinstance(rng, tuple) else rng
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
    ap.add_argument(
        "--reset-cols",
        default="",
        help="Спочатку очистити вміст вказаних колонок (літери або номери, кома). "
             "Напр.: --reset-cols D,E,F. За замовч. — нічого не чистити.",
    )
    args = ap.parse_args()

    client, deployment = build_client()

    # None -> перший аркуш; інакше передаємо рядок як є.
    sheet_arg = args.sheet

    with ExcelBook(args.file, sheet_arg) as book:
        last = book.last_row()
        if args.reset_cols.strip():
            cols_to_clear = [c.strip() for c in args.reset_cols.split(",") if c.strip()]
            for c in cols_to_clear:
                col_ref = c if not c.isdigit() else int(c)
                if isinstance(col_ref, str):
                    rng = f"{col_ref}{args.header_rows + 1}:{col_ref}{last}"
                    book.ws.Range(rng).ClearContents()
                else:
                    book.ws.Range(
                        book.ws.Cells(args.header_rows + 1, col_ref),
                        book.ws.Cells(last, col_ref),
                    ).ClearContents()
                print(f"  очищено колонку {c} (рядки {args.header_rows + 1}..{last})")
            book.save()
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
                # Повторюємо той самий батч до успіху. Усередині call_llm
                # вже є tenacity (~25 хв ретраїв); тут ще додатковий зовнішній
                # цикл на випадок тривалого падіння мережі.
                max_outer_attempts = 5
                outer_attempt = 0
                wrote = False
                while not wrote:
                    outer_attempt += 1
                    try:
                        results = call_llm(client, deployment, rows)
                        book.write_results(rows, results)
                        wrote = True
                    except Exception as e:
                        if outer_attempt >= max_outer_attempts:
                            print(
                                f"[FATAL] Не вдалось обробити рядки "
                                f"{rows[0].excel_row}..{rows[-1].excel_row} "
                                f"після {max_outer_attempts} зовнішніх спроб: {e}. "
                                f"Завершую. Перезапустіть скрипт — він продовжить "
                                f"з цього ж рядка.",
                                file=sys.stderr,
                            )
                            if processed_since_save > 0:
                                book.save()
                            sys.exit(2)
                        backoff = 60 * outer_attempt
                        print(
                            f"[!] Помилка на рядках "
                            f"{rows[0].excel_row}..{rows[-1].excel_row} "
                            f"(зовн. спроба {outer_attempt}/{max_outer_attempts}): "
                            f"{e}. Чекаю {backoff} c і повторюю той самий батч.",
                            file=sys.stderr,
                        )
                        time.sleep(backoff)
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
