import math
import os
import re
import secrets
import datetime
from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from flask import Flask, flash, redirect, render_template, request, send_file, url_for

# Excel экспортын өнгө/хүрээ (вэб сайтын брэнд өнгөтэй нийцүүлсэн)
NAVY = '141B52'
RED = 'C00000'
GREEN = '107C41'
_EDGE = Side(style='thin', color='D3D9EC')
BORDER = Border(left=_EDGE, right=_EDGE, top=_EDGE, bottom=_EDGE)

app = Flask(__name__)

# Production-д SECRET_KEY-г Environment табаас өгнө. Санамсаргүй түлхүүр бол
# сервер дахин асах бүрд сессүүд (flash мессеж) тасардаг.
app.secret_key = os.environ.get('SECRET_KEY') or secrets.token_hex(16)

# Ажлын хавтас (CWD) хаана ч байсан эх файлууд олдохын тулд app.py-гийн хавтаснаас тооцно.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_FILES = {
    'substitutions': os.path.join(BASE_DIR, 'Орлуулах.xlsx'),
    'required': os.path.join(BASE_DIR, 'ЗБНТЖ.xlsx'),
}

STATE = {'substitutions': [], 'required': [], 'groups': [], 'required_codes': set()}

DEFAULT_MTIMES = {}

CODE_RE = re.compile(r'^[A-ZА-ЯӨҮЁ0-9][A-ZА-ЯӨҮЁ0-9./_-]{2,}$', re.IGNORECASE)
DIGIT_RE = re.compile(r'^\d{6,8}$')
NUM_CLEAN_RE = re.compile(r'[,₮% ]')
DATE_RE = re.compile(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})')


def clean(value):
    if value is None:
        return ''
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def code_like(value):
    return bool(CODE_RE.match(clean(value)))


def get(row, index):
    if 0 <= index < len(row):
        return row[index]
    return ''


def to_number(value):
    s = NUM_CLEAN_RE.sub('', clean(value))
    if s == '':
        return 0.0
    try:
        return float(s)
    except ValueError:
        return float('nan')


def js_number_or_zero(value):
    s = clean(value)
    if s == '':
        return 0
    try:
        n = float(s)
    except ValueError:
        return 0
    return 0 if n == 0 else n


def numeric(row):
    values = [n for n in (to_number(v) for v in row) if math.isfinite(n) and n > 0]
    return values[-1] if values else 0


def sales_amount(row):
    n = to_number(get(row, 19))
    return 0 if math.isnan(n) or n == 0 else n


def rows_from_workbook(data):
    workbook = openpyxl.load_workbook(BytesIO(data), data_only=True, read_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    return [[clean(value) for value in row] for row in sheet.iter_rows(values_only=True)]


def parse_groups(rows):
    header = next((row for row in rows if any('дотоод код' in clean(v).lower() for v in row)), None)
    header_index = rows.index(header) if header is not None else -1
    source_code_index = -1
    alternative_code_indexes = []
    if header is not None:
        source_code_index = next((i for i, v in enumerate(header) if clean(v).lower() == 'дотоод код'), -1)
        alternative_code_indexes = [
            i for i, v in enumerate(header)
            if 'орлуулах' in clean(v).lower() and 'дотоод код' in clean(v).lower()
        ]

    groups = []
    for row_index, row in enumerate(rows):
        if source_code_index >= 0:
            code_index = source_code_index
        else:
            code_index = next((i for i, v in enumerate(row) if code_like(v)), -1)
        if code_index < 0:
            continue

        source = {'code': get(row, code_index), 'name': get(row, code_index + 1) or get(row, code_index)}
        alternatives = []
        if alternative_code_indexes:
            candidate_indexes = alternative_code_indexes
        else:
            length = max(0, math.ceil((len(row) - code_index - 2) / 2))
            candidate_indexes = [code_index + 2 + i * 2 for i in range(length)]
        for index in candidate_indexes:
            value = get(row, index)
            if code_like(value):
                alternatives.append({'code': value, 'name': get(row, index + 1) or value})

        if row_index > header_index and source_code_index >= 0 and code_like(source['code']):
            groups.append([source] + alternatives)

    return groups


def expand_products(groups):
    expanded = []
    for group in groups:
        for product in group:
            expanded.append({
                'code': product['code'],
                'name': product['name'],
                'alternatives': [candidate for candidate in group if candidate['code'] != product['code']],
            })

    product_by_code = {}
    alternatives_by_code = {}
    for product in expanded:
        product_by_code[product['code']] = {'code': product['code'], 'name': product['name']}
        alternatives_by_code.setdefault(product['code'], {})

    for product in expanded:
        for alternative in product['alternatives']:
            alternatives_by_code[product['code']][alternative['code']] = alternative
            alternatives_by_code.setdefault(alternative['code'], {})
            alternatives_by_code[alternative['code']][product['code']] = product_by_code[product['code']]

    return [
        {
            'code': code,
            'name': product['name'],
            'alternatives': list(alternatives_by_code.get(code, {}).values()),
        }
        for code, product in product_by_code.items()
    ]


def parse_products(rows):
    return expand_products(parse_groups(rows))


def header_code_column(rows):
    """'Дотоод код' баганын мөр/багана индексийг олно."""
    for row_index, row in enumerate(rows):
        for column, value in enumerate(row):
            if clean(value).lower() == 'дотоод код':
                return row_index, column
    return -1, -1


def stock_columns(rows, header_index):
    """Толгой хэсгээс 'Үлдэгдэл' баганануудыг олно."""
    columns = []
    for row in rows[header_index:header_index + 3]:
        for column, value in enumerate(row):
            if clean(value).lower() == 'үлдэгдэл' and column not in columns:
                columns.append(column)
    return columns


def required_codes_from(rows):
    """ЗБНТЖ файлын 'Дотоод код' баганаас нэр төрлийн жагсаалт гаргана."""
    header_index, code_column = header_code_column(rows)
    if code_column < 0:
        return set()
    return {
        code for code in (clean(get(row, code_column)) for row in rows[header_index + 1:])
        if code and code_like(code)
    }


def warehouse_stock_codes(rows):
    """Агуулахын тайлангаас үлдэгдэлтэй (>0) нэр төрлийн кодуудыг гаргана.

    Тайланд үлдэгдэлгүй бараа ч жагсдаг тул зөвхөн жагсаалтад байгаагаар нь
    тооцвол хангалт үргэлж 100% гарна.
    """
    header_index, code_column = header_code_column(rows)
    if code_column < 0:
        return set()
    columns = stock_columns(rows, header_index)

    codes = set()
    for row in rows[header_index + 1:]:
        code = clean(get(row, code_column))
        if not code or not code_like(code):
            continue
        if not columns:
            codes.add(code)
            continue
        for column in columns:
            value = to_number(get(row, column))
            if math.isfinite(value) and value > 0:
                codes.add(code)
                break
    return codes


PERIOD_RE = re.compile(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})')
PERIOD_LABELS = ('тайлант хугацаа', 'тайлангийн огноо', 'тайлант үе')


def parse_period_days(rows, default):
    """Тайлангийн 'Тайлант хугацаа: A - B' мөрөөс хоногийн тоог гаргана.

    Толгойд нь тийм мөр байхгүй бол default-ыг буцаана (жишээ нь өмнөх сарын
    борлуулалтын тайланд толгойн мэдээлэл байдаггүй).
    """
    for row in rows:
        for value in row:
            text = clean(value)
            if not any(label in text.lower() for label in PERIOD_LABELS):
                continue
            found = PERIOD_RE.findall(text)
            if len(found) >= 2:
                try:
                    start = datetime.date(int(found[0][0]), int(found[0][1]), int(found[0][2]))
                    end = datetime.date(int(found[-1][0]), int(found[-1][1]), int(found[-1][2]))
                except ValueError:
                    continue
                return max(1, abs((end - start).days) + 1)
    return default


def sales_columns(rows):
    """Борлуулалтын тайлангийн 'Цэвэр дүн' бүлгийн Тоо/Нийт дүн баганыг олно.

    Толгой нь хоёр давхар: дээд мөрөнд 'Цэвэр дүн', доод мөрөнд 'Тоо', 'Нийт дүн'.
    Олдохгүй бол энэ тайлангийн ердийн байрлал (18, 19)-ыг ашиглана.
    """
    header_index, code_column = header_code_column(rows)
    quantity_column, amount_column = 18, 19
    if header_index >= 0:
        for column, value in enumerate(rows[header_index]):
            if clean(value).lower() == 'цэвэр дүн':
                quantity_column, amount_column = column, column + 1
                break
    return header_index, code_column, quantity_column, amount_column


def sales_map(rows):
    """Борлуулалтын тайлангаас код бүрийн цэвэр тоо ширхэг ба үнийн дүнг гаргана."""
    header_index, code_column, quantity_column, amount_column = sales_columns(rows)
    if code_column < 0:
        return {}
    totals = {}
    for row in rows[header_index + 1:]:
        code = clean(get(row, code_column))
        if not code or not code_like(code):
            continue
        entry = totals.setdefault(code, {'units': 0, 'revenue': 0})
        entry['units'] += js_number_or_zero(get(row, quantity_column))
        entry['revenue'] += js_number_or_zero(get(row, amount_column))
    return totals


def sales_price_map(rows):
    """Борлуулалтын тайлангийн 'Худалдах үнэ' (E багана) — код тус бүрээр."""
    header_index, code_column = header_code_column(rows)
    if code_column < 0:
        return {}
    price_column = 4
    for column, value in enumerate(rows[header_index]):
        if 'худалдах үнэ' in clean(value).lower():
            price_column = column
            break

    prices = {}
    for row in rows[header_index + 1:]:
        code = clean(get(row, code_column))
        if not code or not code_like(code) or code in prices:
            continue
        price = js_number_or_zero(get(row, price_column))
        if price:
            prices[code] = price
    return prices


def warehouse_prices(rows):
    """Агуулахын тайлангийн 'Худалдах нэгж үнэ' — код тус бүрээр.

    Хамгийн өргөн хамрах хүрээтэй бодит нэгж үнийн эх сурвалж.
    """
    header_index, code_column = header_code_column(rows)
    if code_column < 0:
        return {}
    price_column = 5  # 'Худалдах нэгж үнэ' — энэ тайлангийн ердийн байрлал
    for column, value in enumerate(rows[header_index]):
        label = clean(value).lower()
        if 'үнэ' in label and 'худалдах' in label:
            price_column = column
            break

    prices = {}
    for row in rows[header_index + 1:]:
        code = clean(get(row, code_column))
        if not code or not code_like(code) or code in prices:
            continue
        price = js_number_or_zero(get(row, price_column))
        if price:
            prices[code] = price
    return prices


def outage_prices(rows):
    """Тасалдлын тайлангийн G багана (индекс 6) = 'Нэгж үнэ'."""
    header_index, code_column = header_code_column(rows)
    if code_column < 0:
        return {}
    prices = {}
    for row in rows[header_index + 1:]:
        code = clean(get(row, code_column))
        if not code or not code_like(code) or code in prices:
            continue
        price = js_number_or_zero(get(row, 6))
        if price:
            prices[code] = price
    return prices


def outage_events(rows):
    """Тасалдлын тайлангаас бүтээгдэхүүн бүрийн алдсан борлуулалтыг нэгтгэнэ.

    Нэг бүтээгдэхүүн хэд хэдэн удаа тасарч болох тул мөр бүрийг тусад нь биш,
    дотоод кодоор нь нэгтгэж тооцно. Тайлангийн 'Нийт' мөр код агуулаагүй тул
    өөрөө шүүгдэж үлдэнэ.
      11 багана = 'Алдсан борлуулалт' (тоо ширхэг)
      12 багана = 'Нийт алдагдсан борлуулалт' (үнийн дүн)
    """
    totals = {}
    names = {}

    def add(code, row):
        entry = totals.setdefault(code, {'units': 0, 'revenue': 0, 'days': 0})
        entry['units'] += js_number_or_zero(get(row, 11))
        entry['revenue'] += js_number_or_zero(get(row, 12))
        entry['days'] += js_number_or_zero(get(row, 7))  # 'Тасарсан хоног'
        if code not in names:
            names[code] = ''

    header_index, code_column = header_code_column(rows)
    if code_column >= 0:
        for row in rows[header_index + 1:]:
            code = clean(get(row, code_column))
            if code and code_like(code):
                add(code, row)
                if not names[code]:
                    names[code] = clean(get(row, code_column + 1))
    else:
        # 'Дотоод код' толгой олдоогүй бол хуучин эвристикт шилжинэ.
        for row in rows:
            code = find_code(row)
            if code:
                add(code, row)
                if not names[code]:
                    names[code] = find_product_name(row, code)

    return totals, names


def find_code(row):
    for value in row:
        if DIGIT_RE.match(clean(value)):
            return clean(value)
    for value in row:
        if code_like(value):
            return clean(value)
    return ''


def find_product_name(row, code):
    try:
        index = row.index(code)
    except ValueError:
        return ''
    return get(row, index + 1) or ''


def analysis_days(rows):
    text = ' '.join(clean(value) for row in rows for value in row)
    dates = []
    for match in DATE_RE.finditer(text):
        try:
            dates.append(datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3))))
        except ValueError:
            continue
    if len(dates) >= 2:
        return max(1, (max(dates) - min(dates)).days + 1)
    days = [n for n in (to_number(get(row, 7)) for row in rows) if math.isfinite(n) and n > 0]
    return max(days) if days else 30


def resolve_file(filename):
    """Файлын нэрийг том/жижиг үсэг харгалзахгүйгээр олно."""
    if os.path.isfile(filename):
        return filename
    folder = os.path.dirname(filename) or '.'
    target = os.path.basename(filename).lower()
    try:
        for entry in os.listdir(folder):
            if entry.lower() == target:
                return os.path.join(folder, entry)
    except OSError:
        pass
    return filename


def load_defaults(force=False):
    """Эх файлуудыг (Орлуулах, ЗБНТЖ) уншина.

    Диск дээрх файл өөрчлөгдсөн үед л дахин уншдаг тул эх файлаа сольсны дараа
    серверийг дахин ажиллуулах шаардлагагүй. Өөрчлөгдсөн бол True буцаана.
    """
    changed = False
    for key, filename in DEFAULT_FILES.items():
        path = resolve_file(filename)
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            mtime = None
        if not force and DEFAULT_MTIMES.get(key) == mtime and STATE[key]:
            continue

        try:
            with open(path, 'rb') as handle:
                rows = rows_from_workbook(handle.read())
            groups = parse_groups(rows)
            STATE[key] = expand_products(groups)
            if key == 'substitutions':
                STATE['groups'] = groups
            if key == 'required':
                STATE['required_codes'] = required_codes_from(rows)
        except Exception:
            STATE[key] = []
            if key == 'substitutions':
                STATE['groups'] = []
            if key == 'required':
                STATE['required_codes'] = set()

        DEFAULT_MTIMES[key] = mtime
        changed = True
    return changed


def analyze(sales_rows, previous_rows, warehouse_rows, outage_rows):
    # ЗБНТЖ-н хангалт: агуулахад үлдэгдэлтэй байгаа ЗБНТЖ нэр төрлийн тоо / ЗБНТЖ-н нийт тоо
    required_codes = STATE['required_codes'] or {item['code'] for item in STATE['required']}
    warehouse_codes = warehouse_stock_codes(warehouse_rows)
    matched_required_codes = required_codes & warehouse_codes
    coverage = min(100, round(len(matched_required_codes) / len(required_codes) * 100)) if required_codes else 0

    outage_map, outage_names = outage_events(outage_rows)
    outage_codes = list(outage_map.keys())

    sale_map = sales_map(sales_rows)
    previous_map = sales_map(previous_rows)
    price_map = outage_prices(outage_rows)
    sale_price_map = sales_price_map(sales_rows)
    warehouse_price_map = warehouse_prices(warehouse_rows)

    def unit_price(product_code):
        """Бодит нэгж үнэ: агуулахын 'Худалдах нэгж үнэ' (5-р багана) тэргүүн ээлжинд.

        Борлуулалтын тайлангийн E багана заримдаа бүх бараанд нэг утга агуулдаг
        тул хамгийн сүүлд, өөр эх сурвалж олдоогүй үед л хэрэглэнэ.
        """
        if product_code in warehouse_price_map:
            return warehouse_price_map[product_code]
        if product_code in price_map:
            return price_map[product_code]
        sale = sale_map.get(product_code, {})
        if sale.get('units'):
            return sale['revenue'] / sale['units']
        return sale_price_map.get(product_code, 0)

    # ӨДБ-г өмнөх сарын тайлангийн хугацаанд хуваан гаргаж, шинжилгээний
    # (борлуулалтын тайлангийн) хоногоор үржүүлнэ.
    previous_days = parse_period_days(previous_rows, 30)
    analysis_period_days = parse_period_days(sales_rows, analysis_days(outage_rows))
    substitutions_by_code = {item['code']: item for item in STATE['substitutions']}

    results = []
    for code in outage_codes[:200]:
        outage_name = outage_names.get(code, '')
        source = substitutions_by_code.get(code)
        source = dict(source) if source else {'code': code, 'name': outage_name or 'Тодорхойгүй нэр төрөл', 'alternatives': []}
        if outage_name:
            source['name'] = outage_name
        alternatives = source['alternatives'] if source['alternatives'] else [{'code': '', 'name': 'Орлуулах мэдээлэл байхгүй'}]

        lost_units = outage_map[code]['units']
        lost_revenue = outage_map[code]['revenue']

        # Орлуулагч бүрээр: (энэ сарын ӨДБ − өмнөх сарын ӨДБ) × тасарсан өдөр.
        # Үнэ нь энэ сарын борлуулалтын тайлангийн 'Худалдах үнэ' (E багана).
        outage_days = outage_map[code]['days']
        alternative_sales = []
        for item in alternatives:
            member = item['code']
            current_daily = sale_map.get(member, {}).get('units', 0) / analysis_period_days
            previous_daily = previous_map.get(member, {}).get('units', 0) / previous_days
            member_units = (current_daily - previous_daily) * outage_days
            member_price = unit_price(member)
            alternative_sales.append({
                'code': member,
                'name': item['name'],
                'units': member_units,
                'revenue': member_units * member_price,
            })

        sold = sum(item['units'] for item in alternative_sales)
        revenue = sum(item['revenue'] for item in alternative_sales)

        results.append({
            'code': code,
            'name': source['name'],
            'alternatives': alternative_sales,
            'outageDays': outage_days,
            'hasAlternatives': any(item['code'] for item in alternative_sales),
            'sold': sold,
            'revenue': revenue,
            'lostUnits': lost_units,
            'lostRevenue': lost_revenue,
            # Бодит тасалдлын дүн (зөрүү) = алдсан − орлуулсан
            'gapRevenue': lost_revenue - revenue,
            'required': code in required_codes,
        })

    # Орлуулах бараатай нэр төрлүүд эхэнд, дотроо орлуулалтын дүнгээр эрэмбэлэгдэнэ.
    results.sort(key=lambda row: (not row['hasAlternatives'], -row['revenue'], -row['lostRevenue']))
    return {
        'coverage': coverage,
        'matched': len(matched_required_codes),
        'required_total': len(required_codes),
        # Нийт дүнг харуулах 200 мөрийн хязгаараас үл хамааран бүтнээр нь тооцно.
        'lost_units_total': sum(item['units'] for item in outage_map.values()),
        'lost_revenue_total': sum(item['revenue'] for item in outage_map.values()),
        'outage_total': len(outage_map),
        'substituted_units_total': sum(row['sold'] for row in results),
        'substituted_revenue_total': sum(row['revenue'] for row in results),
        # Хэдэн нэр төрөлд бодитоор орлуулан борлуулалт хийгдсэн бэ
        'substituted_count': sum(1 for row in results if row['revenue'] > 0),
        # Бүлгүүдэд багтсан орлуулагч бараануудын нийт борлуулалтын дүн
        # (давхардуулахгүйн тулд ялгаатай кодоор нь нэгтгэнэ).
        'group_sales_total': sum(
            sale_map.get(member, {}).get('revenue', 0)
            for member in {
                alt['code']
                for row in results for alt in row['alternatives'] if alt['code']
            }
        ),
        # Бодит тасалдлын дүн (зөрүү) = алдсан нийт − орлуулсан нийт
        'gap_revenue_total': (
            sum(item['revenue'] for item in outage_map.values()) - sum(row['revenue'] for row in results)
        ),
        'previous_days': previous_days,
        'analysis_period_days': analysis_period_days,
        'analysis': results,
    }


def match_rank(item, query):
    """Хайлтын тэмдэгтүүд яг тухайн дарааллаараа хаана таарснаар эрэмбэлнэ.

    Бага тоо = эхэнд. Өөрийнх нь код/нэр таарсан нь орлуулах хувилбараараа
    таарснаас үргэлж түрүүлнэ.
    """
    code = item['code'].lower()
    name = item['name'].lower()
    if code.startswith(query):
        return 0
    if name.startswith(query):
        return 1
    if query in code:
        return 2
    if query in name:
        return 3
    for alt in item['alternatives']:
        if alt['code'].lower().startswith(query) or alt['name'].lower().startswith(query):
            return 4
    for alt in item['alternatives']:
        if query in alt['code'].lower() or query in alt['name'].lower():
            return 5
    return None


def search_substitutions(query):
    query = (query or '').strip().lower()
    if not query:
        return []

    ranked = []
    for item in STATE['substitutions']:
        rank = match_rank(item, query)
        if rank is not None:
            ranked.append((rank, item))
    ranked.sort(key=lambda pair: pair[0])
    return [item for _, item in ranked[:12]]


def money(value):
    return f"₮{round(value or 0):,}"


app.jinja_env.filters['money'] = money

EMPTY_ANALYSIS = {
    'coverage': 0, 'analysis': [], 'matched': 0, 'required_total': 0,
    'lost_units_total': 0, 'lost_revenue_total': 0, 'outage_total': 0,
    'substituted_units_total': 0, 'substituted_revenue_total': 0,
    'substituted_count': 0, 'gap_revenue_total': 0, 'group_sales_total': 0,
    'previous_days': 0, 'analysis_period_days': 0,
}

LAST_ANALYSIS = dict(EMPTY_ANALYSIS)

LOGO_FILE = 'monos-logo.png'


def logo_exists():
    return os.path.isfile(os.path.join(app.static_folder, LOGO_FILE))

DOWNLOAD_HEADER = [
    'Дотоод код', 'Тасалдсан нэр төрөл', 'Алдсан борлуулалтын тоо хэмжээ', 'Алдсан борлуулалтын үнийн дүн',
    'Орлуулсан бүтээгдэхүүн', 'Орлуулсан борлуулалтын тоо хэмжээ', 'Орлуулсан борлуулалтын үнийн дүн',
    'Тасарсан хоног', 'Бодит тасалдлын дүн (зөрүү)',
]


@app.route('/')
def index():
    # Эх файл солигдсон бол өмнөх шинжилгээ хуучирсан тул цэвэрлэнэ.
    if load_defaults():
        LAST_ANALYSIS.update(EMPTY_ANALYSIS)
    query = request.args.get('q', '')
    return render_template(
        'index.html',
        query=query,
        matches=search_substitutions(query),
        groups=STATE['groups'],
        products=STATE['substitutions'],
        logo_exists=logo_exists(),
        substitutions_loaded=bool(STATE['substitutions']),
        coverage=LAST_ANALYSIS['coverage'],
        matched=LAST_ANALYSIS['matched'],
        required_total=LAST_ANALYSIS['required_total'],
        lost_units_total=LAST_ANALYSIS['lost_units_total'],
        lost_revenue_total=LAST_ANALYSIS['lost_revenue_total'],
        outage_total=LAST_ANALYSIS['outage_total'],
        substituted_units_total=LAST_ANALYSIS['substituted_units_total'],
        substituted_revenue_total=LAST_ANALYSIS['substituted_revenue_total'],
        substituted_count=LAST_ANALYSIS['substituted_count'],
        gap_revenue_total=LAST_ANALYSIS['gap_revenue_total'],
        group_sales_total=LAST_ANALYSIS['group_sales_total'],
        previous_days=LAST_ANALYSIS['previous_days'],
        analysis_period_days=LAST_ANALYSIS['analysis_period_days'],
        analysis=LAST_ANALYSIS['analysis'],
    )


@app.route('/analyze', methods=['POST'])
def analyze_route():
    # Боловсруулалт бүрийн өмнө эх файлуудын хамгийн сүүлийн хувилбарыг авна.
    load_defaults()

    def file_rows(field):
        uploaded = request.files.get(field)
        return rows_from_workbook(uploaded.read()) if uploaded else []

    sales_rows = file_rows('sales')
    previous_rows = file_rows('previous')
    warehouse_rows = file_rows('warehouse')
    outage_rows = file_rows('outage')

    outcome = analyze(sales_rows, previous_rows, warehouse_rows, outage_rows)
    results = outcome['analysis']
    LAST_ANALYSIS.update(outcome)
    if not results:
        flash('Тасалдлын тайлангаас боловсруулах мөр олдсонгүй.')
    return redirect(url_for('index'))


@app.route('/clear', methods=['POST'])
def clear_route():
    LAST_ANALYSIS.update(EMPTY_ANALYSIS)
    return redirect(url_for('index'))


@app.route('/download')
def download_route():
    results = LAST_ANALYSIS['analysis']
    if not results:
        flash('Эхлээд шинжилгээ хийнэ үү.')
        return redirect(url_for('index'))

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Шинжилгээ'

    # --- Гарчгийн блок ---
    last_column = len(DOWNLOAD_HEADER)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    title = sheet.cell(row=1, column=1, value='ТАСАЛДСАН НЭР ТӨРЛИЙН ОРЛУУЛАЛТЫН ШИНЖИЛГЭЭ')
    title.font = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
    title.alignment = Alignment(horizontal='center', vertical='center')
    title.fill = PatternFill('solid', fgColor=NAVY)
    sheet.row_dimensions[1].height = 30

    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    subtitle = sheet.cell(row=2, column=1, value=(
        f"Боловсруулсан огноо: {datetime.date.today():%Y-%m-%d}    ·    "
        f"Хэвийн түвшин = орлуулагчдын ӨДБ (өмнөх сар / {LAST_ANALYSIS['previous_days']} хоног) "
        f"× {LAST_ANALYSIS['analysis_period_days']} хоног    ·    "
        f"ЗБНТ хангалт: {LAST_ANALYSIS['coverage']}%    ·    Нийт {len(results)} нэр төрөл"
    ))
    subtitle.font = Font(name='Calibri', size=10, italic=True, color='44506E')
    subtitle.alignment = Alignment(horizontal='center', vertical='center')
    subtitle.fill = PatternFill('solid', fgColor='EEF1FA')
    sheet.row_dimensions[2].height = 22

    # --- Толгой мөр ---
    header_row = 4
    for column, name in enumerate(DOWNLOAD_HEADER, start=1):
        cell = sheet.cell(row=header_row, column=column, value=name)
        cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    sheet.row_dimensions[header_row].height = 46

    # --- Өгөгдлийн мөрүүд ---
    money_columns = {4, 7, 9}
    number_columns = {3, 8}
    for index, row in enumerate(results):
        excel_row = header_row + 1 + index
        values = [
            row['code'],
            row['name'],
            row['lostUnits'],
            row['lostRevenue'],
            '\n'.join(f"{item['code']} {item['name']}" for item in row['alternatives']),
            '\n'.join(f"{item['code']}: {round(item['units'])} ш" for item in row['alternatives']),
            '\n'.join(f"{item['code']}: {round(item['revenue']):,} ₮" for item in row['alternatives']),
            row['outageDays'],
            row['gapRevenue'],
        ]
        stripe = PatternFill('solid', fgColor='F7F9FE') if index % 2 else None
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=excel_row, column=column, value=value)
            cell.font = Font(name='Calibri', size=10)
            cell.border = BORDER
            if stripe:
                cell.fill = stripe
            if column in money_columns:
                cell.number_format = '#,##0 ₮'
                cell.alignment = Alignment(horizontal='right', vertical='top')
            elif column in number_columns:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='top')
            elif column in (5, 6, 7):
                cell.alignment = Alignment(vertical='top', wrap_text=True)
            else:
                cell.alignment = Alignment(vertical='top', wrap_text=True)

        sheet.cell(row=excel_row, column=1).alignment = Alignment(horizontal='center', vertical='top')
        sheet.cell(row=excel_row, column=2).font = Font(name='Calibri', size=10, bold=True)
        if row['required']:
            sheet.cell(row=excel_row, column=2).font = Font(name='Calibri', size=10, bold=True, color=RED)
        # Бодит тасалдлын дүн: эерэг = нөхөгдөөгүй үлдсэн (улаан), сөрөг = илүү нөхсөн (ногоон).
        sheet.cell(row=excel_row, column=9).font = Font(
            name='Calibri', size=10, bold=True,
            color=RED if row['gapRevenue'] > 0 else GREEN,
        )

    # --- Нийт мөр ---
    total_row = header_row + 1 + len(results)
    sheet.cell(row=total_row, column=1, value='НИЙТ')
    sheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    for column in range(1, last_column + 1):
        cell = sheet.cell(row=total_row, column=column)
        cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='2747C7')
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='right' if column > 2 else 'center', vertical='center')
    for column in money_columns | number_columns:
        letter = get_column_letter(column)
        cell = sheet.cell(row=total_row, column=column)
        cell.value = f'=SUM({letter}{header_row + 1}:{letter}{total_row - 1})'
        cell.number_format = '#,##0 ₮' if column in money_columns else '#,##0'
    sheet.row_dimensions[total_row].height = 22

    # --- Багануудын өргөн ---
    widths = [11, 34, 13, 17, 40, 26, 22, 13, 20]
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width

    # --- Шүүлт, царцаалт ---
    sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(last_column)}{total_row - 1}"
    sheet.freeze_panes = sheet.cell(row=header_row + 1, column=3)

    # --- Хуудасны тохиргоо (хэвлэхэд бэлэн) ---
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_title_rows = f'{header_row}:{header_row}'
    sheet.print_options.horizontalCentered = True
    sheet.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    sheet.oddFooter.right.text = 'Хуудас &P / &N'
    sheet.oddFooter.right.size = 9
    sheet.oddFooter.left.text = 'Монос — Орлуулалтын шинжилгээ'
    sheet.oddFooter.left.size = 9

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name='SUBG_боловсруулалт.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


load_defaults()

if __name__ == '__main__':
    # Байршуулсан орчинд 0.0.0.0:8000 (Dokploy-н Container Port) дээр сонсоно.
    app.run(
        host=os.environ.get('HOST', '0.0.0.0'),
        port=int(os.environ.get('PORT', '8000')),
        debug=os.environ.get('DEBUG', '').lower() in ('1', 'true', 'yes'),
    )
