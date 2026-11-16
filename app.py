from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from html import escape
from io import BytesIO, StringIO
import csv
import math
import random
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

import streamlit as st
from rossmann import evaluate_and_forecast, is_rossmann_csv, load_rossmann, promotion_summary, store_contexts, weekday_summary


st.set_page_config(page_title="SalesPulse | Dự báo bán hàng", page_icon=":material/monitoring:", layout="wide")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
DERIVED_REVENUE = "Doanh thu (tự tính: đơn giá × số lượng)"


def _column_number(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference.upper())
    number = 0
    for char in letters.group(0) if letters else "A":
        number = number * 26 + ord(char) - 64
    return number - 1


def _read_shared_strings(book: zipfile.ZipFile) -> list[str]:
    try:
        stream = book.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    values = []
    for _, element in ET.iterparse(stream, events=("end",)):
        if element.tag.endswith("}si"):
            values.append("".join(node.text or "" for node in element.iter() if node.tag.endswith("}t")))
            element.clear()
    stream.close()
    return values


def _worksheet_path(book: zipfile.ZipFile) -> str:
    try:
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        first_sheet = workbook.find("m:sheets/m:sheet", NS)
        rel_id = first_sheet.attrib.get(f"{{{NS['r']}}}id")
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        for item in rels.findall("rel:Relationship", NS):
            if item.attrib.get("Id") == rel_id:
                target = item.attrib["Target"].lstrip("/")
                return target if target.startswith("xl/") else "xl/" + target
    except (ET.ParseError, AttributeError, KeyError, TypeError):
        pass
    return "xl/worksheets/sheet1.xml"


def _xlsx_rows(payload: bytes):
    with zipfile.ZipFile(BytesIO(payload)) as book:
        shared = _read_shared_strings(book)
        with book.open(_worksheet_path(book)) as sheet:
            for _, element in ET.iterparse(sheet, events=("end",)):
                if not element.tag.endswith("}row"):
                    continue
                row = []
                for cell in element:
                    if not cell.tag.endswith("}c"):
                        continue
                    index = _column_number(cell.attrib.get("r", "A1"))
                    while len(row) <= index:
                        row.append("")
                    kind = cell.attrib.get("t")
                    if kind == "inlineStr":
                        value = "".join(part.text or "" for part in cell.iter() if part.tag.endswith("}t"))
                    else:
                        value_node = cell.find("m:v", NS)
                        value = value_node.text if value_node is not None and value_node.text else ""
                        if kind == "s" and value:
                            value = shared[int(value)]
                    row[index] = value
                element.clear()
                yield row


def _parse_date(value: str, date_pattern: str | None = None) -> date | None:
    text = str(value).strip()
    if not text:
        return None
    if date_pattern:
        try:
            return datetime.strptime(text, date_pattern).date()
        except ValueError:
            return None
    try:
        serial = float(text)
        if 20000 < serial < 80000:
            return date(1899, 12, 30) + timedelta(days=serial)
    except ValueError:
        pass
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for pattern in ("%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
                    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%m/%d/%y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def _number(value: str) -> float | None:
    try:
        cleaned = str(value).strip().replace("\u00a0", "").replace(" ", "").replace("$", "").replace("£", "").replace("€", "")
        cleaned = re.sub(r"[\u20a0-\u20cf\u00a2-\u00a5]", "", cleaned)
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1]
        if "," in cleaned and "." in cleaned:
            if cleaned.rfind(",") > cleaned.rfind("."):
                cleaned = cleaned.replace(".", "").replace(",", ".")
            else:
                cleaned = cleaned.replace(",", "")
        elif "," in cleaned:
            trailing_digits = len(cleaned) - cleaned.rfind(",") - 1
            cleaned = cleaned.replace(",", "" if trailing_digits == 3 else ".")
        return float(cleaned)
    except (TypeError, ValueError):
        return None


@st.cache_data(show_spinner=False)
def _read_tabular_upload(payload: bytes, filename: str):
    """Read one CSV/XLSX table. ZIP members are selected before calling this."""
    lower = filename.lower()
    if lower.endswith(".csv"):
        try:
            text = payload.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = payload.decode("cp1252", errors="replace")
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        rows = list(csv.reader(StringIO(text), dialect))
    elif lower.endswith(".xlsx"):
        rows = list(_xlsx_rows(payload))
    else:
        raise ValueError("Định dạng chưa hỗ trợ. Hãy tải lên CSV, XLSX hoặc ZIP chứa CSV/XLSX.")
    if not rows or not any(str(cell).strip() for cell in rows[0]):
        raise ValueError("Không tìm thấy dòng tiêu đề trong tệp.")
    headers = [str(value).strip() or f"Cột {index + 1}" for index, value in enumerate(rows[0])]
    return headers, rows[1:]


@st.cache_data(show_spinner=False)
def _list_upload_tables(payload: bytes, filename: str, max_depth: int = 3):
    """List tabular files in a direct upload, ZIP, or nested Kaggle ZIP."""
    candidates = []

    def visit(content: bytes, current_name: str, display_path: str, depth: int):
        lower = current_name.lower()
        if lower.endswith((".csv", ".xlsx")):
            candidates.append({"label": display_path or current_name,
                               "filename": current_name, "payload": content})
            return
        if not lower.endswith(".zip"):
            return
        if depth > max_depth:
            return
        try:
            with zipfile.ZipFile(BytesIO(content)) as archive:
                entries = sorted(archive.infolist(), key=lambda item: item.filename.lower())
                for entry in entries:
                    if entry.is_dir() or entry.filename.startswith("__MACOSX/"):
                        continue
                    basename = entry.filename.rsplit("/", 1)[-1]
                    if basename.startswith("."):
                        continue
                    if entry.filename.lower().endswith((".csv", ".xlsx", ".zip")):
                        member = archive.read(entry)
                        nested_path = f"{display_path} / {entry.filename}" if display_path else entry.filename
                        visit(member, basename, nested_path, depth + 1)
        except zipfile.BadZipFile:
            raise ValueError(f"Tệp '{display_path or current_name}' có đuôi ZIP nhưng không đọc được.")

    visit(payload, filename, filename, 0)
    return candidates


def _preferred_table_index(candidates):
    """Prefer historical training tables over test/submission/helper files."""
    if not candidates:
        return 0
    def score(item):
        name = item["filename"].lower().rsplit(".", 1)[0]
        points = 0
        if name in {"train", "training", "sales", "transactions", "orders", "data", "retail"}:
            points += 100
        if any(token in name for token in ("train", "sales", "transaction", "order", "retail")):
            points += 30
        if any(token in name for token in ("test", "submission", "sample", "feature", "store", "calendar", "item")):
            points -= 60
        return points
    best = max(range(len(candidates)), key=lambda index: score(candidates[index]))
    return best if score(candidates[best]) > 0 else 0


def _suggest_column(headers, aliases, allow_none=False):
    def normalize(value):
        plain = unicodedata.normalize("NFKD", value.strip().lower())
        plain = "".join(char for char in plain if not unicodedata.combining(char))
        return re.sub(r"[^a-z0-9]+", " ", plain).strip()
    normalized = {normalize(header): header for header in headers}
    for alias in aliases:
        key = normalize(alias)
        if key in normalized:
            return normalized[key]
    # Accept common decorated names such as "Revenue (USD)" or "Sales Amount".
    alias_keys = [normalize(alias) for alias in aliases]
    for header in headers:
        candidate = normalize(header)
        if any(len(key) >= 4 and key in candidate for key in alias_keys):
            return header
    return "(Bỏ qua)" if allow_none else headers[0]


def _make_unique_headers(headers):
    counts = defaultdict(int)
    result = []
    for header in headers:
        counts[header] += 1
        result.append(header if counts[header] == 1 else f"{header} ({counts[header]})")
    return result


@st.cache_data(show_spinner=False)
def _aggregate_uploaded_data(headers, rows, date_column, metric_column, quantity_column, group_column, order_column, filename, date_pattern=None, price_column=None):
    positions = {name: index for index, name in enumerate(headers)}
    metric_is_derived = metric_column == DERIVED_REVENUE
    if date_column not in positions or (not metric_is_derived and metric_column not in positions):
        raise ValueError("Hãy chọn đúng cột ngày và cột chỉ tiêu số cần phân tích.")
    if metric_is_derived and (price_column not in positions or quantity_column not in positions):
        raise ValueError("Để tính doanh thu, hãy chọn cột đơn giá và cột số lượng.")
    daily_metric = defaultdict(float)
    daily_quantity = defaultdict(float)
    group_metric = defaultdict(float)
    group_quantity = defaultdict(float)
    orders = set()
    valid_rows = 0
    quantity_selected = quantity_column != "(Bỏ qua)" and quantity_column in positions and quantity_column != metric_column
    group_selected = group_column != "(Bỏ qua)" and group_column in positions
    order_selected = order_column != "(Bỏ qua)" and order_column in positions

    def cell(row, column):
        index = positions[column]
        return row[index] if index < len(row) else ""

    for row in rows:
        if not row:
            continue
        parsed_day = _parse_date(cell(row, date_column), date_pattern)
        quantity = _number(cell(row, quantity_column)) if quantity_selected else 0.0
        if metric_is_derived:
            unit_price = _number(cell(row, price_column))
            metric = unit_price * quantity if unit_price is not None and quantity is not None else None
        else:
            metric = _number(cell(row, metric_column))
        if parsed_day is None or metric is None or quantity is None:
            continue
        group = str(cell(row, group_column)).strip() if group_selected else "Tất cả dữ liệu"
        group = group or "Không rõ"
        daily_metric[parsed_day] += metric
        daily_quantity[parsed_day] += quantity
        group_metric[group] += metric
        group_quantity[group] += quantity
        if order_selected:
            order_id = str(cell(row, order_column)).strip()
            if order_id:
                orders.add(order_id)
        valid_rows += 1
    if not daily_metric:
        raise ValueError("Không tìm thấy dòng hợp lệ. Kiểm tra ngày và cột chỉ tiêu, hoặc đổi cách ánh xạ cột.")
    return {
        "daily_revenue": {day.isoformat(): value for day, value in daily_metric.items()},
        "daily_quantity": {day.isoformat(): value for day, value in daily_quantity.items()},
        "country_revenue": dict(group_metric), "country_quantity": dict(group_quantity),
        "orders": len(orders) if order_selected else valid_rows, "rows": valid_rows,
        "source": f"Tệp tải lên · {filename}", "market_label": group_column if group_selected else "Nhóm",
        "metric_label": "Doanh thu tính từ đơn giá × số lượng" if metric_is_derived else metric_column,
        "quantity_label": quantity_column if quantity_selected else "Số lượng bán",
        "has_quantity": quantity_selected, "currency": "",
    }


def _period_start(day: date, frequency: str) -> date:
    if frequency == "Tuần":
        return day + timedelta(days=(6 - day.weekday()) % 7)
    if frequency == "Tháng":
        return date(day.year, day.month, 1)
    return day


def _next_period(day: date, frequency: str) -> date:
    if frequency == "Ngày":
        return day + timedelta(days=1)
    if frequency == "Tuần":
        return day + timedelta(days=7)
    return date(day.year + (day.month == 12), day.month % 12 + 1, 1)


def make_series(data: dict, frequency: str, target: str) -> tuple[list[date], list[float]]:
    field = "daily_revenue" if target == "Revenue" else "daily_quantity"
    sales_by_day = {date.fromisoformat(key): value for key, value in data[field].items()}
    totals = defaultdict(float)
    for day, value in sales_by_day.items():
        totals[_period_start(day, frequency)] += value
    start, end = min(totals), max(totals)
    dates, values = [], []
    current = start
    while current <= end:
        dates.append(current)
        values.append(totals.get(current, 0.0))
        current = _next_period(current, frequency)
    return dates, values


def _lags(frequency: str) -> tuple[list[int], list[int]]:
    return {"Ngày": ([1, 7, 14, 28], [7, 28]),
            "Tuần": ([1, 4, 8, 13], [4, 13]),
            "Tháng": ([1, 2, 3], [2, 3])}[frequency]


def _features(history: list[float], day: date, step: int, frequency: str) -> list[float]:
    lag_values, windows = _lags(frequency)
    result = [history[-lag] for lag in lag_values]
    result.extend(sum(history[-window:]) / window for window in windows)
    result.extend([step / 100.0, math.sin(day.weekday() * 2 * math.pi / 7),
                   math.cos(day.weekday() * 2 * math.pi / 7),
                   math.sin(day.month * 2 * math.pi / 12), math.cos(day.month * 2 * math.pi / 12)])
    return result


def _training_data(dates: list[date], values: list[float], frequency: str):
    lags, _ = _lags(frequency)
    first = max(lags)
    x = [_features(values[:i], dates[i], i, frequency) for i in range(first, len(values))]
    y = values[first:]
    return x, y, first


def _solve(matrix: list[list[float]], vector: list[float]) -> list[float]:
    size = len(vector)
    augmented = [matrix[row][:] + [vector[row]] for row in range(size)]
    for column in range(size):
        pivot = max(range(column, size), key=lambda row: abs(augmented[row][column]))
        if abs(augmented[pivot][column]) < 1e-12:
            augmented[pivot][column] += 1e-8
        augmented[column], augmented[pivot] = augmented[pivot], augmented[column]
        divisor = augmented[column][column]
        for j in range(column, size + 1):
            augmented[column][j] /= divisor
        for row in range(size):
            if row == column:
                continue
            factor = augmented[row][column]
            for j in range(column, size + 1):
                augmented[row][j] -= factor * augmented[column][j]
    return [augmented[row][-1] for row in range(size)]


def _fit_linear(x: list[list[float]], y: list[float]):
    width = len(x[0])
    means = [sum(row[col] for row in x) / len(x) for col in range(width)]
    scales = []
    for col in range(width):
        variance = sum((row[col] - means[col]) ** 2 for row in x) / len(x)
        scales.append(math.sqrt(variance) or 1.0)
    matrix = [[0.0] * (width + 1) for _ in range(width + 1)]
    vector = [0.0] * (width + 1)
    for row, target in zip(x, y):
        scaled = [1.0] + [(row[col] - means[col]) / scales[col] for col in range(width)]
        for i, a in enumerate(scaled):
            vector[i] += a * target
            for j, b in enumerate(scaled):
                matrix[i][j] += a * b
    for col in range(1, width + 1):
        matrix[col][col] += 1e-8
    return {"means": means, "scales": scales, "weights": _solve(matrix, vector)}


def _predict_linear(model, row: list[float]) -> float:
    scaled = [1.0] + [(value - model["means"][i]) / model["scales"][i] for i, value in enumerate(row)]
    return sum(a * b for a, b in zip(scaled, model["weights"]))


def _fit_tree(x: list[list[float]], y: list[float], rng: random.Random, depth: int = 0):
    mean = sum(y) / len(y)
    if depth >= 5 or len(y) < 12:
        return ("leaf", mean)
    width = len(x[0])
    feature_ids = rng.sample(range(width), max(2, int(math.sqrt(width))))
    best = None
    min_leaf = 5
    for feature in feature_ids:
        ordered = sorted((row[feature], target) for row, target in zip(x, y))
        count = len(ordered)
        stride = max(1, count // 12)
        total = sum(target for _, target in ordered)
        total_sq = sum(target * target for _, target in ordered)
        left_sum = left_sq = 0.0
        for index, (value, target) in enumerate(ordered[:-1], start=1):
            left_sum += target
            left_sq += target * target
            if index < min_leaf or count - index < min_leaf or index % stride:
                continue
            if value == ordered[index][0]:
                continue
            right_sum, right_sq = total - left_sum, total_sq - left_sq
            error = left_sq - left_sum * left_sum / index + right_sq - right_sum * right_sum / (count - index)
            if best is None or error < best[0]:
                best = (error, feature, (value + ordered[index][0]) / 2)
    if best is None:
        return ("leaf", mean)
    _, feature, threshold = best
    left_x, left_y, right_x, right_y = [], [], [], []
    for row, target in zip(x, y):
        if row[feature] <= threshold:
            left_x.append(row); left_y.append(target)
        else:
            right_x.append(row); right_y.append(target)
    if not left_y or not right_y:
        return ("leaf", mean)
    return ("node", feature, threshold, _fit_tree(left_x, left_y, rng, depth + 1), _fit_tree(right_x, right_y, rng, depth + 1))


def _tree_predict(tree, row: list[float]) -> float:
    while tree[0] == "node":
        tree = tree[3] if row[tree[1]] <= tree[2] else tree[4]
    return tree[1]


def _fit_forest(x: list[list[float]], y: list[float]):
    rng = random.Random(42)
    trees = []
    for _ in range(32):
        sample = [rng.randrange(len(y)) for _ in range(len(y))]
        trees.append(_fit_tree([x[i] for i in sample], [y[i] for i in sample], rng))
    return trees


def _predict(model, row: list[float], model_name: str) -> float:
    if model_name == "Linear Regression":
        return _predict_linear(model, row)
    return sum(_tree_predict(tree, row) for tree in model) / len(model)


def _forecast(model, history_values: list[float], future_dates: list[date], frequency: str, model_name: str):
    history = history_values[:]
    predictions = []
    for day in future_dates:
        row = _features(history, day, len(history), frequency)
        prediction = max(0.0, _predict(model, row, model_name))
        predictions.append(prediction)
        history.append(prediction)
    return predictions


@st.cache_data(show_spinner=False)
def train_forecast(dates: list[date], values: list[float], frequency: str, horizon: int, model_name: str):
    x, y, first = _training_data(dates, values, frequency)
    if len(x) < 8:
        raise ValueError("Cần thêm dữ liệu lịch sử. Hãy chọn Ngày/Tuần hoặc tải tệp có thời gian dài hơn.")
    holdout = min(max(3, horizon), max(3, len(values) // 5))
    split = len(values) - holdout
    cutoff = split - first
    x_train, y_train = x[:cutoff], y[:cutoff]
    if len(x_train) < 5:
        raise ValueError("Dữ liệu quá ngắn để kiểm tra mô hình theo thời gian.")
    if model_name == "Linear Regression":
        validation_model = _fit_linear(x_train, y_train)
    else:
        validation_model = _fit_forest(x_train, y_train)
    validation_dates = dates[split:]
    validation_predictions = _forecast(validation_model, values[:split], validation_dates, frequency, model_name)
    actual = values[split:]
    errors = [abs(a - p) for a, p in zip(actual, validation_predictions)]
    mae = sum(errors) / len(errors)
    rmse = math.sqrt(sum((a - p) ** 2 for a, p in zip(actual, validation_predictions)) / len(errors))
    denominator = sum(abs(value) for value in actual)
    wape = sum(errors) / denominator if denominator else 0.0
    if model_name == "Linear Regression":
        final_model = _fit_linear(x, y)
    else:
        final_model = _fit_forest(x, y)
    future_dates = []
    day = _next_period(dates[-1], frequency)
    for _ in range(horizon):
        future_dates.append(day)
        day = _next_period(day, frequency)
    future_predictions = _forecast(final_model, values, future_dates, frequency, model_name)
    return validation_dates, validation_predictions, future_dates, future_predictions, {"mae": mae, "rmse": rmse, "wape": wape, "holdout": holdout}


def _money(value: float) -> str:
    return f"£{value:,.0f}"


def _line_chart(dates, actual, validation_dates, validation_predictions, future_dates, future_predictions, target, money_format=_money, value_label=None):
    width, height, left, right, top, bottom = 920, 350, 68, 20, 24, 48
    visible_start = max(0, len(dates) - 90)
    chart_dates = dates[visible_start:] + future_dates
    values = actual[visible_start:] + validation_predictions + future_predictions
    low, high = min([0.0] + values), max([1.0] + values)
    span = high - low or 1
    def point(index, value):
        x = left + index * (width - left - right) / max(1, len(chart_dates) - 1)
        y = top + (high - value) * (height - top - bottom) / span
        return x, y
    def make_path(start_index, items):
        return " ".join(("M" if offset == 0 else "L") + f"{point(start_index + offset, value)[0]:.1f},{point(start_index + offset, value)[1]:.1f}" for offset, value in enumerate(items))
    actual_path = make_path(0, actual[visible_start:])
    validation_start = next((i for i, day in enumerate(chart_dates) if day in validation_dates), None)
    validation_path = make_path(validation_start, validation_predictions) if validation_start is not None else ""
    future_start = len(dates[visible_start:]) - 1
    future_path = make_path(future_start, [actual[-1]] + future_predictions)
    grids = []
    for fraction in (0, .25, .5, .75, 1):
        y = top + fraction * (height - top - bottom)
        label = high - fraction * span
        text = money_format(label) if target == "Revenue" else f"{label:,.0f}"
        grids.append(f"<line x1='{left}' y1='{y:.1f}' x2='{width-right}' y2='{y:.1f}' stroke='#e9eef5'/><text x='{left-8}' y='{y+4:.1f}' text-anchor='end' fill='#76859a' font-size='11'>{escape(text)}</text>")
    xlabels = []
    label_indexes = sorted(set([0, len(chart_dates)//4, len(chart_dates)//2, 3*len(chart_dates)//4, len(chart_dates)-1]))
    for index in label_indexes:
        if chart_dates:
            x, _ = point(index, low)
            xlabels.append(f"<text x='{x:.1f}' y='{height-12}' text-anchor='middle' fill='#76859a' font-size='11'>{chart_dates[index].strftime('%d %b %y')}</text>")
    unit = value_label or ("giá trị" if target == "Revenue" else "số lượng")
    return f"""<div style='font:12px Arial,sans-serif;background:white;border-radius:12px;overflow:auto'>
    <div style='display:flex;gap:18px;padding:8px 10px;color:#52647d'><span style='color:#3278df'>● Thực tế</span><span style='color:#f07843'>● Dự báo kiểm tra / tương lai</span></div>
    <svg viewBox='0 0 {width} {height}' width='100%' role='img' aria-label='Biểu đồ doanh số thực tế và dự báo'>
    {''.join(grids)}<path d='{actual_path}' fill='none' stroke='#3278df' stroke-width='2.6'/>
    <path d='{validation_path}' fill='none' stroke='#ed8150' stroke-width='2.3' stroke-dasharray='5 4'/>
    <path d='{future_path}' fill='none' stroke='#ed8150' stroke-width='3'/>{''.join(xlabels)}</svg>
    <div style='padding:0 12px 10px;color:#718096'>Đơn vị: {unit}</div></div>"""


def _country_chart(data, money_format=_money):
    entries = sorted(data["country_revenue"].items(), key=lambda item: item[1], reverse=True)[:10]
    max_value = max((value for _, value in entries), default=1)
    rows = []
    for country, value in entries:
        width = max(1, value / max_value * 100)
        rows.append(f"<div style='display:grid;grid-template-columns:95px 1fr 72px;align-items:center;gap:8px;margin:10px 0;font:12px Arial;color:#45566f'><span style='overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>{escape(country)}</span><div style='height:12px;background:#edf3fa;border-radius:8px'><div style='height:12px;width:{width:.1f}%;background:#4d91e8;border-radius:8px'></div></div><span style='text-align:right'>{escape(money_format(value))}</span></div>")
    return "<div style='padding:6px 4px;background:white'>" + "".join(rows) + "</div>"


def _rossmann_table(headers, rows):
    head = "".join(f"<th style='padding:9px;text-align:left'>{escape(str(item))}</th>" for item in headers)
    body = "".join("<tr>" + "".join(
        f"<td style='padding:9px;border-bottom:1px solid #e6ebf2'>{escape(str(value))}</td>" for value in row
    ) + "</tr>" for row in rows)
    return ("<table style='border-collapse:collapse;width:100%;background:white;font:13px Arial,sans-serif;color:#1c2e4a'>"
            f"<thead><tr style='background:#eef3fa'>{head}</tr></thead><tbody>{body}</tbody></table>")


def _weekday_month_matrix(daily_values):
    names = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]
    months = ["T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8", "T9", "T10", "T11", "T12"]
    cells = [[0.0] * 12 for _ in range(7)]
    for raw_day, value in daily_values.items():
        day = date.fromisoformat(raw_day) if isinstance(raw_day, str) else raw_day
        cells[day.weekday()][day.month - 1] += value
    return names, months, cells


def _echarts_option(labels, values, chart_type, *, secondary=None, matrix=None, title=""):
    """Build a JSON-only ECharts option so charts render natively in Streamlit."""
    entries = [(str(label), float(value)) for label, value in zip(labels, values)
               if math.isfinite(float(value))]
    if len(entries) > 180:
        indexes = [round(i * (len(entries) - 1) / 179) for i in range(180)]
        entries = [entries[index] for index in indexes]
        if secondary is not None:
            secondary = [secondary[index] for index in indexes if index < len(secondary)]
    shown_labels = [label for label, _ in entries]
    shown_values = [value for _, value in entries]
    option = {
        "title": {"text": title, "left": 18, "top": 8, "textStyle": {"fontSize": 15, "fontWeight": 600}},
        "tooltip": {"trigger": "axis", "valueFormatter": "%,.0f"},
        "grid": {"left": 68, "right": 28, "top": 58, "bottom": 58, "containLabel": True},
        "animation": False,
    }
    if chart_type in ("Đường", "Vùng"):
        option.update({"xAxis": {"type": "category", "data": shown_labels, "boundaryGap": False},
                       "yAxis": {"type": "value", "name": "Giá trị"},
                       "dataZoom": [{"type": "inside"}, {"type": "slider", "height": 18, "bottom": 8}],
                       "series": [{"name": title, "type": "line", "smooth": True, "showSymbol": len(shown_values) < 50,
                                   "data": shown_values, "areaStyle": {"opacity": .2} if chart_type == "Vùng" else None,
                                   "lineStyle": {"width": 3}, "itemStyle": {"color": "#2584e8"}}]})
    elif chart_type in ("Cột", "Thanh ngang"):
        horizontal = chart_type == "Thanh ngang"
        shown = sorted(zip(shown_labels, shown_values), key=lambda row: row[1], reverse=True)[:30]
        cat, vals = [item[0] for item in shown], [item[1] for item in shown]
        option.update({"grid": {"left": 130 if horizontal else 68, "right": 34, "top": 58, "bottom": 48, "containLabel": True},
                       "xAxis": {"type": "value" if horizontal else "category", **({} if horizontal else {"data": cat})},
                       "yAxis": {"type": "category" if horizontal else "value", **({"data": cat, "inverse": True} if horizontal else {})},
                       "series": [{"type": "bar", "data": vals, "itemStyle": {"color": "#2584e8", "borderRadius": [4, 4, 0, 0] if not horizontal else [0, 4, 4, 0]}, "barMaxWidth": 38}]})
    elif chart_type == "Tròn":
        items = sorted(zip(shown_labels, shown_values), key=lambda row: row[1], reverse=True)[:10]
        option.update({"tooltip": {"trigger": "item", "valueFormatter": "%,.0f"}, "legend": {"orient": "vertical", "left": "55%", "top": "middle"},
                       "series": [{"type": "pie", "radius": ["38%", "72%"], "center": ["32%", "55%"], "data": [{"name": label, "value": value} for label, value in items],
                                   "label": {"formatter": "{b}: {d}%"}}]})
    elif chart_type == "Treemap":
        items = sorted(zip(shown_labels, shown_values), key=lambda row: row[1], reverse=True)[:30]
        option.update({"tooltip": {"valueFormatter": "%,.0f"}, "series": [{"type": "treemap", "roam": False, "breadcrumb": {"show": False}, "data": [{"name": label, "value": value} for label, value in items],
                                    "label": {"show": True, "formatter": "{b}"}, "upperLabel": {"show": False}}]})
    elif chart_type == "Phân tán":
        secondary = secondary or list(range(len(shown_values)))
        points = [[float(x), float(y)] for x, y in zip(secondary, shown_values)]
        option.update({"tooltip": {"trigger": "item", "formatter": "X: {@[0]}<br/>Y: {@[1]}"},
                       "xAxis": {"type": "value", "name": "Số lượng / khách hàng", "scale": True},
                       "yAxis": {"type": "value", "name": "Doanh thu", "scale": True},
                       "series": [{"type": "scatter", "data": points, "symbolSize": 8, "itemStyle": {"color": "#2584e8", "opacity": .65}}]})
    elif chart_type == "Histogram":
        if not shown_values:
            return option
        low, high = min(shown_values), max(shown_values)
        span = high - low or 1
        bin_count = min(14, max(5, int(math.sqrt(len(shown_values)))))
        counts = [0] * bin_count
        for value in shown_values:
            counts[min(bin_count - 1, int((value - low) / span * bin_count))] += 1
        bin_labels = [f"{low + i*span/bin_count:,.0f}–{low + (i+1)*span/bin_count:,.0f}" for i in range(bin_count)]
        option.update({"tooltip": {"trigger": "axis"}, "xAxis": {"type": "category", "data": bin_labels, "axisLabel": {"rotate": 25}},
                       "yAxis": {"type": "value", "name": "Số kỳ"}, "series": [{"type": "bar", "data": counts, "itemStyle": {"color": "#16a085"}}]})
    elif chart_type == "Heatmap" and matrix:
        row_labels, col_labels, cells = matrix
        flat_values = [value for row in cells for value in row]
        max_value = max(flat_values, default=1) or 1
        heat = [[j, i, value] for i, row in enumerate(cells) for j, value in enumerate(row)]
        option.update({"tooltip": {"position": "top", "valueFormatter": "%,.0f"}, "grid": {"left": 90, "right": 30, "top": 58, "bottom": 36},
                       "xAxis": {"type": "category", "data": col_labels, "splitArea": {"show": True}},
                       "yAxis": {"type": "category", "data": row_labels, "splitArea": {"show": True}},
                       "visualMap": {"min": 0, "max": max_value, "calculable": True, "orient": "horizontal", "left": "center", "bottom": 0},
                       "series": [{"type": "heatmap", "data": heat, "label": {"show": False}, "emphasis": {"itemStyle": {"shadowBlur": 8, "shadowColor": "rgba(0,0,0,0.35)"}}}]})
    return option


def _rossmann_line_chart(dates, values, validation_dates, validation_predictions,
                         forecast_dates, forecast_predictions, visible_days=180):
    width, height, left, right, top, bottom = 940, 360, 78, 22, 28, 48
    start = max(0, len(dates) - visible_days)
    history_dates, history_values = dates[start:], values[start:]
    all_dates = history_dates + forecast_dates
    scale_values = history_values + validation_predictions + forecast_predictions
    low, high = min([0.0] + scale_values), max([1.0] + scale_values)
    span = high - low or 1

    def point(index, value):
        x = left + index * (width - left - right) / max(1, len(all_dates) - 1)
        y = top + (high - value) * (height - top - bottom) / span
        return x, y

    def path(index, series):
        if not series:
            return ""
        return " ".join(("M" if offset == 0 else "L") +
                        f"{point(index + offset, value)[0]:.1f},{point(index + offset, value)[1]:.1f}"
                        for offset, value in enumerate(series))

    actual_path = path(0, history_values)
    validation_index = next((i for i, day in enumerate(all_dates) if validation_dates and day == validation_dates[0]), None)
    validation_path = path(validation_index, validation_predictions) if validation_index is not None else ""
    forecast_path = path(len(history_dates) - 1, [values[-1]] + forecast_predictions)
    grid = []
    for part in (0, .25, .5, .75, 1):
        y = top + part * (height - top - bottom)
        label = high - part * span
        grid.append(f"<line x1='{left}' y1='{y:.1f}' x2='{width-right}' y2='{y:.1f}' stroke='#e7edf4'/>"
                    f"<text x='{left-9}' y='{y+4:.1f}' text-anchor='end' fill='#718096' font-size='11'>{label:,.0f}</text>")
    xlabels = []
    for index in sorted(set([0, len(all_dates)//3, 2*len(all_dates)//3, len(all_dates)-1])):
        if all_dates:
            x, _ = point(index, low)
            xlabels.append(f"<text x='{x:.1f}' y='{height-12}' text-anchor='middle' fill='#718096' font-size='11'>{all_dates[index].strftime('%d %b %y')}</text>")
    return f"""<div style='font:12px Arial,sans-serif;background:white;border-radius:12px'>
    <div style='display:flex;gap:16px;padding:8px 10px;color:#53647a'><span style='color:#2563eb'>● Thực tế</span><span style='color:#f97316'>● Dự báo kiểm tra / tương lai</span></div>
    <svg viewBox='0 0 {width} {height}' width='100%' role='img' aria-label='Doanh thu thực tế và dự báo theo ngày'>
    {''.join(grid)}<path d='{actual_path}' fill='none' stroke='#2563eb' stroke-width='2.5'/>
    <path d='{validation_path}' fill='none' stroke='#f97316' stroke-width='2.2' stroke-dasharray='5 4'/>
    <path d='{forecast_path}' fill='none' stroke='#f97316' stroke-width='3'/>{''.join(xlabels)}</svg>
    <div style='padding:0 12px 10px;color:#718096'>Doanh thu theo đơn vị trong dataset</div></div>"""


def _rossmann_store_bars(summary, highest=True):
    ordered = sorted(summary.items(), key=lambda pair: pair[1]["sales"], reverse=highest)[:10]
    high = max((row["sales"] for _, row in ordered), default=1)
    elements = []
    for store_id, row in ordered:
        width = max(1, row["sales"] / high * 100)
        elements.append(f"<div style='display:grid;grid-template-columns:70px 1fr 85px;gap:8px;align-items:center;margin:8px 0;font:12px Arial;color:#43536a'>"
                        f"<span>Cửa hàng {escape(store_id)}</span><div style='height:12px;background:#edf2f8;border-radius:8px'><div style='height:12px;width:{width:.1f}%;background:#2563eb;border-radius:8px'></div></div>"
                        f"<span style='text-align:right'>{row['sales']:,.0f}</span></div>")
    return "<div style='background:white;padding:4px'>" + "".join(elements) + "</div>"


def _importance_chart(importance):
    strongest = sorted(importance, key=lambda item: item["importance"], reverse=True)[:8]
    max_value = max((item["importance"] for item in strongest), default=1) or 1
    rows = []
    for item in strongest:
        width = item["importance"] / max_value * 100
        rows.append(f"<div style='display:grid;grid-template-columns:190px 1fr 58px;gap:8px;align-items:center;margin:8px 0;font:12px Arial;color:#43536a'>"
                    f"<span>{escape(item['feature'])}</span><div style='height:11px;background:#edf2f8;border-radius:7px'><div style='height:11px;width:{width:.1f}%;background:#2563eb;border-radius:7px'></div></div>"
                    f"<span style='text-align:right'>{item['importance']:.1f}%</span></div>")
    return "<div style='background:white;padding:6px'>" + "".join(rows) + "</div>"


def _weekday_bars(profile):
    high = max((item["average"] for item in profile), default=1) or 1
    columns = []
    for item in profile:
        height = max(3, item["average"] / high * 100)
        columns.append(f"<div style='flex:1;text-align:center;font:11px Arial;color:#53647a'>"
                       f"<div style='height:130px;display:flex;align-items:end;justify-content:center'><div title='{item['average']:,.0f}' style='width:65%;height:{height:.1f}%;background:#3b82f6;border-radius:5px 5px 0 0'></div></div>"
                       f"<div style='margin-top:7px'>{item['day']}</div><div style='color:#8592a5'>{item['average']:,.0f}</div></div>")
    return "<div style='display:flex;gap:8px;align-items:end;background:white;padding:10px'>" + "".join(columns) + "</div>"


@st.cache_data(show_spinner=False)
def _cached_rossmann(train_payload: bytes, stores_payload: bytes | None, test_payload: bytes | None):
    return load_rossmann(train_payload, stores_payload, test_payload)


@st.cache_data(show_spinner=False)
def _cached_evaluation(dates, values, contexts, horizon, future_dates, future_contexts):
    return evaluate_and_forecast(dates, values, contexts, horizon, future_dates, future_contexts)


def _get_uploaded_rossmann(uploads):
    csv_files = [(item.name, item.getvalue()) for item in uploads if item.name.lower().endswith(".csv")]
    zip_files = [(item.name, item.getvalue()) for item in uploads if item.name.lower().endswith(".zip")]
    extracted = []
    for _, payload in zip_files:
        try:
            with zipfile.ZipFile(BytesIO(payload)) as archive:
                for name in archive.namelist():
                    if name.lower().endswith(".csv"):
                        extracted.append((name.rsplit("/", 1)[-1], archive.read(name)))
        except zipfile.BadZipFile:
            continue
    files = csv_files + extracted
    train = next((payload for _, payload in files if is_rossmann_csv(payload)), None)
    if train is None:
        return None
    store_file = next((payload for name, payload in files
                       if "store" in name.lower() and not is_rossmann_csv(payload)), None)
    test_file = None
    for name, payload in files:
        if payload == train or payload == store_file:
            continue
        try:
            headers = {str(key).strip().lower() for key in (csv.DictReader(StringIO(payload.decode("utf-8-sig", errors="replace"))).fieldnames or [])}
        except Exception:
            continue
        if {"store", "date"}.issubset(headers) and "sales" not in headers:
            test_file = payload
            break
    return train, store_file, test_file


def _rossmann_shared_filters(data):
    dates = data["dates"]
    metadata = data["metadata"]
    with st.sidebar:
        st.divider()
        st.markdown("**Bộ lọc dashboard**")
        states = sorted({row.get("State", "") for row in metadata.values()} - {""})
        cities = sorted({row.get("City", "") for row in metadata.values()} - {""})
        store_types = sorted({row.get("StoreType", "") for row in metadata.values()} - {""})
        selected_states = st.multiselect("Bang / khu vực", states, default=[])
        selected_cities = st.multiselect("Thành phố", cities, default=[])
        selected_types = st.pills("Loại cửa hàng", store_types, selection_mode="multi") or []
        selected_dates = st.date_input("Khoảng ngày phân tích", value=(dates[0], dates[-1]),
                                       min_value=dates[0], max_value=dates[-1])
        frequency = st.segmented_control("Gom xu hướng", ["Ngày", "Tuần", "Tháng"], default="Tuần") or "Tuần"
        filtered_stores = []
        for store_id in data["stores"]:
            row = metadata.get(store_id, {})
            if selected_states and row.get("State", "") not in selected_states:
                continue
            if selected_cities and row.get("City", "") not in selected_cities:
                continue
            if selected_types and row.get("StoreType", "") not in selected_types:
                continue
            filtered_stores.append(store_id)
        filtered_stores.sort(key=lambda value: (not value.isdigit(), int(value) if value.isdigit() else value))
        if not filtered_stores:
            st.warning("Bộ lọc hiện không có cửa hàng nào.")
            st.stop()
        if st.session_state.get("dashboard_store_id") not in filtered_stores:
            st.session_state["dashboard_store_id"] = filtered_stores[0]
        selected_store = st.selectbox("Cửa hàng chi tiết / dự báo", filtered_stores,
                                      key="dashboard_store_id",
                                      format_func=lambda value: f"{value} · {metadata.get(value, {}).get('City', '—')} · {metadata.get(value, {}).get('StoreType', '—')}")
    if isinstance(selected_dates, (tuple, list)):
        start_day = selected_dates[0] if selected_dates else dates[0]
        end_day = selected_dates[-1] if len(selected_dates) > 1 else start_day
    else:
        start_day = end_day = selected_dates
    start_index = max(0, (start_day - dates[0]).days)
    end_index = min(len(dates) - 1, (end_day - dates[0]).days)
    if end_index < start_index:
        start_index, end_index = end_index, start_index
    return {"stores": filtered_stores, "primary": selected_store, "start": start_index,
            "end": end_index, "start_date": dates[start_index], "end_date": dates[end_index],
            "frequency": frequency}


def _rossmann_totals(data, store_ids, start, end, field="sales"):
    end = min(end, len(data["dates"]) - 1)
    start = max(0, start)
    return {store_id: sum(data["stores"][store_id][field][start:end + 1]) for store_id in store_ids}


def _rossmann_group_series(data, store_ids, start, end, frequency, field="sales"):
    totals = defaultdict(float)
    dates = data["dates"]
    for index in range(start, end + 1):
        day = dates[index]
        period = _period_start(day, frequency)
        totals[period] += sum(data["stores"][store_id][field][index] for store_id in store_ids)
    result_dates = sorted(totals)
    return result_dates, [totals[day] for day in result_dates]


def _rossmann_group_profile(data, store_ids, start, end, field="sales"):
    labels = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
    buckets = [[] for _ in range(7)]
    for store_id in store_ids:
        store = data["stores"][store_id]
        for index in range(start, min(end + 1, len(store[field]))):
            if store["open"][index]:
                buckets[data["dates"][index].weekday()].append(store[field][index])
    return [{"day": labels[i], "average": sum(values) / len(values) if values else 0.0,
             "count": len(values)} for i, values in enumerate(buckets)]


def _rossmann_group_promo(data, store_ids, start, end):
    promo_values, regular_values, holiday_values, ordinary_values = [], [], [], []
    for store_id in store_ids:
        store = data["stores"][store_id]
        for index in range(start, min(end + 1, len(store["sales"]))):
            if not store["open"][index]:
                continue
            sales = store["sales"][index]
            (promo_values if store["promo"][index] else regular_values).append(sales)
            is_holiday = store["state_holiday"][index] or store["school_holiday"][index]
            (holiday_values if is_holiday else ordinary_values).append(sales)
    promo_avg = sum(promo_values) / len(promo_values) if promo_values else 0.0
    regular_avg = sum(regular_values) / len(regular_values) if regular_values else 0.0
    uplift = (promo_avg / regular_avg - 1) * 100 if regular_avg else 0.0
    holiday_avg = sum(holiday_values) / len(holiday_values) if holiday_values else 0.0
    ordinary_avg = sum(ordinary_values) / len(ordinary_values) if ordinary_values else 0.0
    return {"promo_avg": promo_avg, "regular_avg": regular_avg, "uplift": uplift,
            "promo_count": len(promo_values), "regular_count": len(regular_values),
            "holiday_avg": holiday_avg, "ordinary_avg": ordinary_avg,
            "holiday_count": len(holiday_values), "ordinary_count": len(ordinary_values)}


def _page_header(title, subtitle):
    st.title(title)
    st.caption(subtitle)


def _st_navigation(pages):
    navigation = st.navigation(pages, position="top")
    navigation.run()


def _render_rossmann(data):
    filters = _rossmann_shared_filters(data)
    with st.sidebar:
        st.divider()
        horizon = st.selectbox("Kỳ dự báo", [14, 28, 42], index=2, format_func=lambda value: f"{value} ngày")
        model_choice = st.selectbox("Mô hình dự báo", ["Tự chọn theo RMSPE", "Seasonal baseline", "Linear Regression", "Random Forest"])
        history_window = st.select_slider("Lịch sử trên biểu đồ", ["90 ngày", "180 ngày", "Toàn bộ"], value="180 ngày")
        st.caption("Bộ lọc khu vực áp dụng cho các trang phân tích; dự báo được tính đến ngày cuối của train.csv.")

    dates = data["dates"]
    store_ids = filters["stores"]
    store_id = filters["primary"]
    start, end = filters["start"], filters["end"]
    start_date, end_date = filters["start_date"], filters["end_date"]
    selected_store = data["stores"][store_id]
    selected_summary = data["summary"][store_id]
    filtered_totals = _rossmann_totals(data, store_ids, start, end)
    group_dates, group_sales = _rossmann_group_series(data, store_ids, start, end, filters["frequency"])
    group_customers_dates, group_customers = _rossmann_group_series(data, store_ids, start, end, filters["frequency"], "customers")
    group_promo = _rossmann_group_promo(data, store_ids, start, end)
    group_weekday = _rossmann_group_profile(data, store_ids, start, end)
    active_ids = [sid for sid in store_ids if filtered_totals.get(sid, 0) > 0]
    mean_day_sales = sum(filtered_totals.values()) / max(1, end - start + 1)
    ranked = sorted(filtered_totals.items(), key=lambda item: item[1], reverse=True)

    contexts = store_contexts(selected_store)
    weekday_open = {}
    for weekday in range(7):
        observed = [selected_store["open"][i] for i, day in enumerate(dates) if day.weekday() == weekday]
        weekday_open[weekday] = int(sum(observed) >= len(observed) / 2) if observed else 1
    horizon = min(horizon, max(1, len(dates) - 7))
    final_train_date = dates[-1]
    test_dates = [day for day in data["test_dates"] if day > final_train_date]
    future_dates = test_dates[:horizon]
    while len(future_dates) < horizon:
        next_day = future_dates[-1] + timedelta(days=1) if future_dates else final_train_date + timedelta(days=1)
        future_dates.append(next_day)
    future_contexts = [data["future_context"].get((store_id, day),
                       {"open": weekday_open[day.weekday()], "promo": 0, "state_holiday": 0, "school_holiday": 0})
                       for day in future_dates]
    evaluation = _cached_evaluation(dates, selected_store["sales"], contexts, horizon, future_dates, future_contexts)
    results = evaluation["results"]
    best_model = min(results, key=lambda name: results[name]["metrics"]["rmspe"])
    selected_model = best_model if model_choice == "Tự chọn theo RMSPE" else model_choice
    selected = results[selected_model]
    store_info = data["metadata"].get(store_id, {})
    store_label = f"Cửa hàng {store_id} · {store_info.get('City', '—')} · loại {store_info.get('StoreType', '—')}"
    visible_days = {"90 ngày": 90, "180 ngày": 180, "Toàn bộ": len(dates)}[history_window]

    def overview_page():
        _page_header("SalesPulse · Tổng quan", "Bức tranh bán hàng cho phạm vi bộ lọc đang chọn.")
        st.caption(f"Nguồn: Kaggle Rossmann Store Sales · {start_date:%d/%m/%Y} – {end_date:%d/%m/%Y} · {len(store_ids):,} cửa hàng")
        with st.container(horizontal=True):
            st.metric("Doanh thu trong kỳ", f"{sum(filtered_totals.values()):,.0f}", f"{len(active_ids):,} cửa hàng có bán", border=True)
            st.metric("Doanh thu bình quân / ngày", f"{mean_day_sales:,.0f}", filters["frequency"], border=True)
            st.metric(f"Dự báo {horizon} ngày · {store_id}", f"{sum(selected['forecast']):,.0f}", selected_model, border=True)
            st.metric("RMSPE kiểm tra", f"{selected['metrics']['rmspe']:.1f}%", f"WAPE {selected['metrics']['wape']:.1f}%", border=True)
        left, right = st.columns([1.6, 1])
        with left, st.container(border=True):
            st.subheader(f"Xu hướng doanh thu · {filters['frequency'].lower()}")
            st.html(_rossmann_line_chart(group_dates, group_sales, [], [], [], [], min(len(group_dates), visible_days)))
        with right, st.container(border=True):
            st.subheader("Cửa hàng theo doanh thu")
            st.html(_rossmann_store_bars({sid: {"sales": value} for sid, value in ranked[:10]}, highest=True))
            st.caption("Xếp hạng trong khoảng ngày đã lọc.")
        a, b = st.columns(2)
        with a, st.container(border=True):
            st.subheader("Doanh thu trung bình theo thứ")
            st.html(_weekday_bars(group_weekday))
        with b, st.container(border=True):
            st.subheader("Khuyến mãi · mô tả")
            st.metric("Ngày có khuyến mãi", f"{group_promo['promo_avg']:,.0f}", f"{group_promo['promo_count']:,} ngày cửa hàng mở")
            st.metric("Ngày thường", f"{group_promo['regular_avg']:,.0f}", f"{group_promo['regular_count']:,} ngày cửa hàng mở")
            st.caption(f"Chênh lệch trung bình: {group_promo['uplift']:+.1f}%. Đây là so sánh mô tả, không chứng minh quan hệ nhân quả.")

    def trends_page():
        _page_header("Xu hướng", "Chọn chỉ tiêu và mức gộp để xem hành vi bán hàng theo thời gian.")
        target = st.segmented_control("Chỉ tiêu", ["Doanh thu", "Khách hàng"], default="Doanh thu", key="trends_metric") or "Doanh thu"
        series_dates, series_values = (group_dates, group_sales) if target == "Doanh thu" else (group_customers_dates, group_customers)
        with st.container(border=True):
            st.subheader(f"{target} theo {filters['frequency'].lower()}")
            st.html(_rossmann_line_chart(series_dates, series_values, [], [], [], [], len(series_dates)))
        if series_values:
            period = max(1, min(7 if filters["frequency"] == "Ngày" else 4, len(series_values)))
            prior = sum(series_values[-period * 2:-period]) / max(1, len(series_values[-period * 2:-period]))
            latest = sum(series_values[-period:]) / max(1, len(series_values[-period:]))
            change = (latest / prior - 1) * 100 if prior else 0
            x, y, z = st.columns(3)
            x.metric("Kỳ gần nhất", f"{series_values[-1]:,.0f}")
            y.metric("Bình quân kỳ gần đây", f"{latest:,.0f}")
            z.metric("So với kỳ trước", f"{change:+.1f}%")
        rows = [[day.isoformat(), f"{value:,.0f}"] for day, value in zip(series_dates, series_values)]
        st.subheader("Bảng xu hướng")
        st.html(_rossmann_table(["Ngày bắt đầu kỳ", target], rows[-100:]))

    def chart_explorer_page():
        _page_header("Khám phá biểu đồ", "Chọn góc nhìn và loại biểu đồ để tự tạo một lát cắt phân tích.")
        view = st.selectbox("Bạn muốn phân tích", ["Xu hướng theo thời gian", "So sánh cửa hàng", "Phân bố doanh thu ngày", "Doanh thu và khách hàng", "Heatmap thứ × tháng"])
        target_label = "Doanh thu" if view == "Doanh thu và khách hàng" else (st.segmented_control("Chỉ tiêu", ["Doanh thu", "Khách hàng"], default="Doanh thu", key="chart_target") or "Doanh thu")
        target_field = "sales" if target_label == "Doanh thu" else "customers"
        chart_type = None
        labels, values = [], []
        secondary, matrix = None, None
        table_rows = []
        if view == "Xu hướng theo thời gian":
            labels, values = _rossmann_group_series(data, store_ids, start, end, filters["frequency"], target_field)
            labels = [day.strftime("%Y-%m-%d") for day in labels]
            chart_type = st.selectbox("Kiểu biểu đồ", ["Đường", "Vùng", "Cột", "Bảng"])
            table_rows = [[label, f"{value:,.0f}"] for label, value in zip(labels, values)]
        elif view == "So sánh cửa hàng":
            totals = _rossmann_totals(data, store_ids, start, end, target_field)
            ordered = sorted(totals.items(), key=lambda item: item[1], reverse=True)[:20]
            labels, values = [f"Cửa hàng {sid}" for sid, _ in ordered], [value for _, value in ordered]
            chart_type = st.selectbox("Kiểu biểu đồ", ["Thanh ngang", "Cột", "Tròn", "Treemap", "Bảng"])
            table_rows = [[label, f"{value:,.0f}"] for label, value in zip(labels, values)]
        elif view == "Phân bố doanh thu ngày":
            labels, values = [str(i + 1) for i in range(start, end + 1)], [sum(data["stores"][sid]["sales"][i] for sid in store_ids) for i in range(start, end + 1)]
            chart_type = "Histogram"
        elif view == "Doanh thu và khách hàng":
            labels = [day.isoformat() for day in dates[start:end + 1]]
            values = [sum(data["stores"][sid]["sales"][i] for sid in store_ids) for i in range(start, end + 1)]
            secondary = [sum(data["stores"][sid]["customers"][i] for sid in store_ids) for i in range(start, end + 1)]
            if len(labels) > 60:
                stride = math.ceil(len(labels) / 60)
                labels, values, secondary = labels[::stride], values[::stride], secondary[::stride]
            chart_type = "Phân tán"
        else:
            weekday_labels = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
            month_labels = [f"T{i}" for i in range(1, 13)]
            cells = [[0.0] * 12 for _ in range(7)]
            for i in range(start, end + 1):
                day = dates[i]
                cells[day.weekday()][day.month - 1] += sum(data["stores"][sid][target_field][i] for sid in store_ids)
            matrix = weekday_labels, month_labels, cells
            labels, values = [name for name in weekday_labels], [sum(row) for row in cells]
            chart_type = "Heatmap"
        st.caption("Biểu đồ được cập nhật theo bộ lọc khu vực, cửa hàng và khoảng ngày ở thanh bên.")
        if chart_type == "Bảng":
            st.html(_rossmann_table(["Nhóm / kỳ", target_label], table_rows))
        else:
            with st.container(border=True):
                st.echarts_chart(_echarts_option(labels, values, chart_type, secondary=secondary, matrix=matrix, title=f"{view} · {target_label}"),
                                height=430, alt=f"{view} theo {target_label.lower()}")
        if view == "Doanh thu và khách hàng":
            st.caption("Trục ngang là số khách hàng; trục dọc là doanh thu. Mỗi chấm đại diện một ngày trong khoảng lọc.")
        elif view == "Phân bố doanh thu ngày":
            st.caption("Histogram gom doanh thu ngày thành các khoảng giá trị; số phía dưới mỗi cột là số ngày trong khoảng.")

    def stores_page():
        _page_header("Phân tích cửa hàng", "So sánh cửa hàng trong khu vực và ngày được chọn ở thanh bên.")
        store_totals_all = _rossmann_totals(data, store_ids, start, end)
        sorted_totals = sorted(store_totals_all.items(), key=lambda item: item[1], reverse=True)
        total_rank = 1 + sum(value > store_totals_all.get(store_id, 0) for value in store_totals_all.values())
        profile = weekday_summary(selected_store, dates[start:end + 1])
        with st.container(horizontal=True):
            st.metric("Doanh thu trong kỳ", f"{store_totals_all.get(store_id, 0):,.0f}", store_label, border=True)
            st.metric("Thứ hạng trong phạm vi lọc", f"{total_rank}/{len(sorted_totals)}", "theo doanh thu", border=True)
            st.metric("Khách hàng lịch sử", f"{selected_summary['customers']:,.0f}", f"{selected_summary['open_days']:,} ngày mở", border=True)
        chart_dates = dates[start:end + 1]
        chart_values = selected_store["sales"][start:end + 1]
        with st.container(border=True):
            st.subheader(f"Lịch sử · {store_label}")
            st.html(_rossmann_line_chart(chart_dates, chart_values, [], [], [], [], len(chart_dates)))
        a, b = st.columns([1, 1.3])
        with a, st.container(border=True):
            st.subheader("Bán hàng theo thứ")
            st.html(_weekday_bars(profile))
        with b, st.container(border=True):
            st.subheader("Top cửa hàng trong phạm vi lọc")
            st.html(_rossmann_table(["Hạng", "Cửa hàng", "Thành phố", "Loại", "Doanh thu"], [
                [i + 1, sid, data["metadata"].get(sid, {}).get("City", "—"), data["metadata"].get(sid, {}).get("StoreType", "—"), f"{value:,.0f}"]
                for i, (sid, value) in enumerate(sorted_totals[:20])]))

    def promotion_page():
        _page_header("Khuyến mãi & lịch bán", "So sánh ngày promo, ngày lễ và nhịp bán theo thứ trong bộ lọc hiện tại.")
        a, b, c = st.columns(3)
        a.metric("Bình quân ngày promo", f"{group_promo['promo_avg']:,.0f}", f"{group_promo['promo_count']:,} quan sát")
        b.metric("Bình quân ngày thường", f"{group_promo['regular_avg']:,.0f}", f"{group_promo['regular_count']:,} quan sát")
        c.metric("Chênh lệch trung bình", f"{group_promo['uplift']:+.1f}%", "mô tả, không nhân quả")
        a, b = st.columns(2)
        with a, st.container(border=True):
            st.subheader("Doanh thu bình quân theo thứ")
            st.html(_weekday_bars(group_weekday))
        with b, st.container(border=True):
            st.subheader("Ngày lễ và ngày thường")
            st.html(_rossmann_table(["Nhóm ngày", "Doanh thu bình quân", "Số quan sát"], [
                ["Có ngày lễ", f"{group_promo['holiday_avg']:,.0f}", f"{group_promo['holiday_count']:,}"],
                ["Ngày thường", f"{group_promo['ordinary_avg']:,.0f}", f"{group_promo['ordinary_count']:,}"]]))
        st.info("So sánh này chưa kiểm soát khác biệt giữa cửa hàng, mùa vụ và thời điểm triển khai; không nên diễn giải là tác động nhân quả của khuyến mãi.")

    def forecast_page():
        _page_header("Dự báo & đánh giá mô hình", f"Dự báo {horizon} ngày cho {store_label}.")
        if not data["test_dates"]:
            st.info("Đã dùng lịch ngày trong tuần để tạo dự báo. Nạp thêm test.csv để bổ sung promo và ngày lễ tương lai.", icon=":material/info:")
        selected_result = results[selected_model]
        scores = selected_result["metrics"]
        with st.container(horizontal=True):
            st.metric("Tổng doanh thu dự báo", f"{sum(selected_result['forecast']):,.0f}", selected_model, border=True)
            st.metric("RMSPE", f"{scores['rmspe']:.2f}%", f"WAPE {scores['wape']:.2f}%", border=True)
            st.metric("MAE", f"{scores['mae']:,.0f}", f"RMSE {scores['rmse']:,.0f}", border=True)
            st.metric("Holdout theo thời gian", f"{evaluation['holdout']} ngày", "không xáo trộn dữ liệu", border=True)
        with st.container(border=True):
            st.subheader("Thực tế, kiểm tra và tương lai")
            st.html(_rossmann_line_chart(dates, selected_store["sales"], evaluation["validation_dates"],
                                         selected_result["validation"], future_dates, selected_result["forecast"], visible_days))
        score_rows = []
        for name, result in sorted(results.items(), key=lambda item: item[1]["metrics"]["rmspe"]):
            metric = result["metrics"]
            score_rows.append([name, f"{metric['rmspe']:.2f}%", f"{metric['wape']:.2f}%", f"{metric['mae']:,.0f}", f"{metric['rmse']:,.0f}"])
        st.subheader("So sánh mô hình trên kỳ holdout")
        st.html(_rossmann_table(["Mô hình", "RMSPE ↓", "WAPE", "MAE", "RMSE"], score_rows))
        st.caption(f"Tự chọn: {best_model}, có RMSPE thấp nhất trên kỳ holdout. RMSPE có thể nhạy với ngày doanh thu thấp.")
        if "importance" in selected_result:
            st.subheader("Biến mô hình sử dụng nhiều")
            st.html(_importance_chart(selected_result["importance"]))
        st.subheader("Chi tiết dự báo")
        forecast_rows = [[day.isoformat(), f"{value:,.2f}"] for day, value in zip(future_dates, selected_result["forecast"])]
        st.html(_rossmann_table(["Ngày", "Sales dự báo"], forecast_rows))
        csv_lines = ["Store,Date,PredictedSales"] + [f"{store_id},{day.isoformat()},{value:.2f}" for day, value in zip(future_dates, selected_result["forecast"])]
        st.download_button("Tải dự báo CSV", data="\ufeff" + "\n".join(csv_lines), file_name=f"rossmann_store_{store_id}_forecast.csv", mime="text/csv", icon=":material/download:")

    def data_page():
        _page_header("Dữ liệu & phương pháp", "Thông tin nguồn, độ phủ và hướng dẫn đọc kết quả.")
        st.markdown(f"- **Nguồn:** Kaggle Rossmann Store Sales.\n- **Dữ liệu huấn luyện:** {data['rows']:,} dòng, {len(data['stores']):,} cửa hàng.\n- **Thời gian:** {dates[0]:%d/%m/%Y} – {dates[-1]:%d/%m/%Y}.\n- **Phạm vi lọc:** {len(store_ids):,} cửa hàng, {start_date:%d/%m/%Y} – {end_date:%d/%m/%Y}.\n- **Cửa hàng dự báo:** {store_label}.")
        st.subheader("Các mô hình")
        st.markdown("- **Seasonal baseline:** lấy doanh thu cùng ngày trong tuần gần nhất.\n- **Linear Regression:** dùng độ trễ, trung bình trượt, ngày trong tuần/tháng, trạng thái mở cửa, promo và ngày lễ.\n- **Random Forest:** mô hình cây rừng hồi quy cài đặt bằng Python cơ bản để chạy trong môi trường hiện tại.")
        st.subheader("Giới hạn khi diễn giải")
        st.markdown("Dữ liệu Rossmann là dữ liệu bán lẻ tại Đức và doanh thu trong bộ dữ liệu gốc không phải đơn vị tiền tệ được chuẩn hóa. Dự báo là ước lượng từ lịch sử; chất lượng tương lai có thể khác kỳ holdout. Bộ lọc ngày dùng cho phân tích lịch sử; mô hình dự báo được huấn luyện trên toàn bộ chuỗi của cửa hàng đã chọn.")
        if data["test_dates"]:
            st.success(f"Đã đọc lịch test.csv đến {data['test_dates'][-1]:%d/%m/%Y}.")
        else:
            st.warning("Chưa có test.csv; lịch promo/ngày lễ tương lai được đặt mặc định.")

    pages = [
        st.Page(overview_page, title="Tổng quan", icon=":material/dashboard:", default=True),
        st.Page(trends_page, title="Xu hướng", icon=":material/trending_up:"),
        st.Page(chart_explorer_page, title="Khám phá biểu đồ", icon=":material/insert_chart:"),
        st.Page(stores_page, title="Cửa hàng", icon=":material/storefront:"),
        st.Page(promotion_page, title="Khuyến mãi", icon=":material/percent:"),
        st.Page(forecast_page, title="Dự báo", icon=":material/query_stats:"),
        st.Page(data_page, title="Dữ liệu & phương pháp", icon=":material/database:")]
    _st_navigation(pages)


def _render_uploaded_pages(data, frequency, horizon, model_name, target):
    dates, values = make_series(data, frequency, target)
    validation_dates, validation_predictions, future_dates, future_predictions, metrics = train_forecast(
        dates, values, frequency, horizon, model_name)
    total_revenue = sum(data["country_revenue"].values())
    total_quantity = sum(data["country_quantity"].values())
    selected_market_values = data["country_revenue"] if target == "Revenue" else data["country_quantity"]
    countries = sorted(selected_market_values.items(), key=lambda item: item[1], reverse=True)
    target_label = data.get("metric_label", "Chỉ tiêu") if target == "Revenue" else data.get("quantity_label", "Số lượng bán")
    target_total = sum(values)
    prediction_total = sum(future_predictions)
    money = target == "Revenue"
    currency = data.get("currency", "")
    money_format = lambda value: f"{currency}{value:,.0f}"
    format_value = money_format if money else lambda value: f"{value:,.0f}"

    def overview_page():
        _page_header("SalesPulse · Tổng quan", f"Phân tích file {data.get('source', 'đã tải lên')}.")
        st.info(f"Đang phân tích {data.get('rows', 0):,} dòng hợp lệ. Chỉ tiêu chính: **{data.get('metric_label', 'Chỉ tiêu')}**.", icon=":material/info:")
        with st.container(horizontal=True):
            st.metric(f"Tổng {target_label} lịch sử", format_value(target_total), border=True)
            if data.get("has_quantity") and target != "Quantity":
                st.metric(data.get("quantity_label", "Số lượng bán"), f"{total_quantity:,.0f}", f"{data['orders']:,} đơn/dòng", border=True)
            else:
                st.metric("Đơn hàng / dòng hợp lệ", f"{data['orders']:,}", border=True)
            st.metric(f"Dự báo {horizon} kỳ", format_value(prediction_total), model_name, border=True)
            st.metric("Độ khớp kiểm tra (1 − WAPE)", f"{max(0.0, 1 - metrics['wape']):.1%}", f"MAE {format_value(metrics['mae'])}", border=True)
        with st.container(border=True):
            st.subheader(f"{target_label} thực tế và dự báo")
            st.html(_line_chart(dates, values, validation_dates, validation_predictions, future_dates, future_predictions, target, money_format, target_label))
        a, b = st.columns(2)
        with a, st.container(border=True):
            st.subheader(f"Top {data.get('market_label', 'nhóm').lower()} theo {target_label}")
            chart_data = dict(data)
            chart_data["country_revenue"] = selected_market_values
            st.html(_country_chart(chart_data, format_value))
        with b, st.container(border=True):
            st.subheader("Chất lượng kiểm tra")
            st.metric("WAPE", f"{metrics['wape']:.1%}")
            st.metric("RMSE", format_value(metrics["rmse"]))
            st.caption(f"Đánh giá trên {metrics['holdout']} kỳ cuối theo thời gian.")

    def trends_page():
        _page_header("Xu hướng", "Xem lịch sử bán hàng và các kỳ gần nhất.")
        st.caption(f"Chỉ tiêu: {target_label} · gộp theo {frequency.lower()}")
        with st.container(border=True):
            st.html(_line_chart(dates, values, validation_dates, validation_predictions, future_dates, future_predictions, target, money_format, target_label))
        rows = [[day.isoformat(), format_value(value)] for day, value in zip(dates, values)]
        st.subheader("Bảng số liệu theo kỳ")
        st.html(_rossmann_table(["Ngày bắt đầu kỳ", target_label], rows[-100:]))

    def chart_explorer_page():
        _page_header("Khám phá biểu đồ", "Chọn loại phân tích và cách trực quan hóa, giống bộ chọn biểu đồ trong Tableau.")
        relation_label = f"{data.get('metric_label', 'Chỉ tiêu')} × {data.get('quantity_label', 'Số lượng')}"
        views = ["Xu hướng theo thời gian", "So sánh theo thị trường", "Phân bố theo kỳ", "Heatmap thứ × tháng"]
        if data.get("has_quantity"):
            views.insert(3, relation_label)
        view = st.selectbox("Bạn muốn phân tích", views)
        metric_choices = [data.get("metric_label", "Chỉ tiêu")]
        if data.get("has_quantity") and data.get("quantity_label") != data.get("metric_label"):
            metric_choices.append(data.get("quantity_label", "Số lượng bán"))
        selected_target_label = st.segmented_control("Chỉ tiêu", metric_choices, default=metric_choices[0], key="viz_target") or metric_choices[0]
        target_label = data.get("metric_label", "Chỉ tiêu") if view == relation_label else selected_target_label
        target_key = "Revenue" if target_label == data.get("metric_label", "Chỉ tiêu") else "Quantity"
        target_daily = data["daily_revenue"] if target_key == "Revenue" else data["daily_quantity"]
        target_market = data["country_revenue"] if target_key == "Revenue" else data["country_quantity"]
        chart_type, secondary, matrix = None, None, None
        labels, values, table_rows = [], [], []
        if view == "Xu hướng theo thời gian":
            chart_dates, chart_values = make_series(data, frequency, target_key)
            labels, values = [day.isoformat() for day in chart_dates], chart_values
            chart_type = st.selectbox("Kiểu biểu đồ", ["Đường", "Vùng", "Cột", "Bảng"])
            table_rows = [[label, format_value(value)] for label, value in zip(labels, values)]
        elif view == "So sánh theo thị trường":
            sorted_markets = sorted(target_market.items(), key=lambda item: item[1], reverse=True)[:20]
            labels, values = [item[0] for item in sorted_markets], [item[1] for item in sorted_markets]
            chart_type = st.selectbox("Kiểu biểu đồ", ["Thanh ngang", "Cột", "Tròn", "Treemap", "Bảng"])
            table_rows = [[label, format_value(value)] for label, value in zip(labels, values)]
        elif view == "Phân bố theo kỳ":
            chart_dates, chart_values = make_series(data, frequency, target_key)
            labels, values = [day.isoformat() for day in chart_dates], chart_values
            chart_type = "Histogram"
        elif view == relation_label:
            revenue = data["daily_revenue"]
            quantity = data["daily_quantity"]
            dates_scatter = sorted(set(revenue) & set(quantity))
            labels = dates_scatter
            values = [revenue[day] for day in dates_scatter]
            secondary = [quantity[day] for day in dates_scatter]
            chart_type = "Phân tán"
        else:
            matrix = _weekday_month_matrix(target_daily)
            labels, values = matrix[0], [sum(row) for row in matrix[2]]
            chart_type = "Heatmap"
        st.caption(f"Nguồn: {data.get('source', 'dữ liệu mẫu')} · biểu đồ cập nhật theo chỉ tiêu và mức gộp ở thanh bên.")
        if chart_type == "Bảng":
            st.html(_rossmann_table(["Thị trường / kỳ", target_label], table_rows))
        else:
            with st.container(border=True):
                st.echarts_chart(_echarts_option(labels, values, chart_type, secondary=secondary, matrix=matrix, title=f"{view} · {target_label}"),
                                height=430, alt=f"{view} theo {target_label.lower()}")
        if view == relation_label:
            st.caption(f"Mỗi chấm là một ngày; trục ngang là {data.get('quantity_label', 'số lượng')} và trục dọc là {data.get('metric_label', 'chỉ tiêu')}.")
        elif view == "Phân bố theo kỳ":
            st.caption("Histogram gom các kỳ có giá trị gần nhau; nhãn dưới cột là số kỳ trong mỗi khoảng.")

    def markets_page():
        _page_header("Phân tích nhóm", f"So sánh {target_label.lower()} theo {data.get('market_label', 'nhóm').lower()}.")
        market_values = data["country_revenue"] if target == "Revenue" else data["country_quantity"]
        a, b = st.columns(2)
        with a, st.container(border=True):
            st.subheader(f"Top 10 {data.get('market_label', 'Nhóm').lower()} · {target_label}")
            chart_data = dict(data)
            chart_data["country_revenue"] = market_values
            st.html(_country_chart(chart_data, format_value))
        with b, st.container(border=True):
            st.subheader("Bảng xếp hạng")
            market_rows = []
            for index, (country, amount) in enumerate(sorted(market_values.items(), key=lambda item: item[1], reverse=True)[:30], 1):
                market_rows.append([index, country, format_value(amount)])
            st.html(_rossmann_table(["Hạng", data.get("market_label", "Nhóm"), target_label], market_rows))

    def forecast_page():
        _page_header("Dự báo", f"{model_name} · {horizon} kỳ · {frequency.lower()}.")
        with st.container(horizontal=True):
            st.metric(f"Tổng dự báo {target_label.lower()}", format_value(prediction_total), border=True)
            st.metric("WAPE kiểm tra", f"{metrics['wape']:.1%}", f"Holdout {metrics['holdout']} kỳ", border=True)
            st.metric("MAE", format_value(metrics["mae"]), border=True)
            st.metric("RMSE", format_value(metrics["rmse"]), border=True)
        with st.container(border=True):
            st.subheader("Thực tế, kiểm tra và dự báo")
            st.html(_line_chart(dates, values, validation_dates, validation_predictions, future_dates, future_predictions, target, money_format, target_label))
        rows = [[day.isoformat(), format_value(value)] for day, value in zip(future_dates, future_predictions)]
        st.html(_rossmann_table(["Ngày bắt đầu kỳ", f"{target_label} dự báo"], rows))
        csv_lines = ["Date,Forecast"] + [f"{day.isoformat()},{value:.2f}" for day, value in zip(future_dates, future_predictions)]
        st.download_button("Tải dự báo CSV", data="\ufeff" + "\n".join(csv_lines), file_name="sales_forecast.csv", mime="text/csv", icon=":material/download:")
        st.caption(f"Đây là dự báo từ {data.get('source', 'dữ liệu mẫu')} và mô hình đã chọn. Dùng dữ liệu phù hợp với nghiệp vụ trước khi đưa ra quyết định kinh doanh.")

    def data_page():
        _page_header("Dữ liệu & phương pháp", "Nguồn, cách tổng hợp và cách đánh giá mô hình.")
        st.markdown(f"- **Nguồn:** {data.get('source', 'tệp tải lên')}.\n- **Dòng hợp lệ:** {data['rows']:,}.\n- **Tổng {target_label}:** {format_value(target_total)}.\n- **Tổng {data.get('quantity_label', 'số lượng')}:** {total_quantity:,.0f}.\n- **Số nhóm:** {len(selected_market_values):,}.\n- **Chỉ tiêu dự báo:** {target_label}; gộp theo {frequency.lower()}.")
        st.subheader("Quy trình đánh giá")
        st.markdown(f"Mô hình **{model_name}** được kiểm tra trên {metrics['holdout']} kỳ cuối theo thời gian. WAPE tính tổng sai số tuyệt đối chia tổng giá trị thực tế; MAE và RMSE dùng cùng đơn vị với chỉ tiêu.")
        st.warning("Đánh giá dùng các kỳ cuối trong file đã tải lên. Dự báo ngoại suy có thể kém tin cậy khi tương lai có xu hướng hoặc mùa vụ chưa xuất hiện trong lịch sử.")

    _st_navigation([
        st.Page(overview_page, title="Tổng quan", icon=":material/dashboard:", default=True),
        st.Page(trends_page, title="Xu hướng", icon=":material/trending_up:"),
        st.Page(chart_explorer_page, title="Khám phá biểu đồ", icon=":material/insert_chart:"),
        st.Page(markets_page, title="Thị trường", icon=":material/public:"),
        st.Page(forecast_page, title="Dự báo", icon=":material/query_stats:"),
        st.Page(data_page, title="Dữ liệu & phương pháp", icon=":material/database:")])


def main():
    with st.sidebar:
        st.markdown("## :material/monitoring: SalesPulse")
        st.caption("Dự báo doanh thu và hỗ trợ quyết định bán lẻ")
        uploads = st.file_uploader("Tải dữ liệu bán hàng", type=["csv", "zip", "xlsx"],
                                  accept_multiple_files=True,
                                  help="CSV/XLSX hoặc ZIP, kể cả ZIP lồng nhau. Với bộ nhiều bảng, hãy chọn bảng có cột ngày và doanh số/số lượng; Rossmann được nhận diện tự động.")
        st.caption("Không chọn sẵn bộ dữ liệu. File của bạn sẽ là nguồn phân tích.")

    uploads = uploads or []
    rossmann_files = _get_uploaded_rossmann(uploads)
    use_rossmann_mode = bool(rossmann_files)
    if rossmann_files and len(uploads) > 1:
        with st.sidebar:
            data_mode = st.selectbox("Chế độ dữ liệu", ["Rossmann chuyên sâu", "Bảng chung"],
                                     key="upload_data_mode",
                                     help="Chọn Rossmann để dùng dashboard chuyên biệt, hoặc Bảng chung để chọn một bảng khác đã tải lên.")
        use_rossmann_mode = data_mode == "Rossmann chuyên sâu"
    if use_rossmann_mode:
        train_payload, stores_payload, test_payload = rossmann_files
        try:
            with st.spinner("Đang đọc Rossmann và chuẩn bị các mô hình…"):
                data = _cached_rossmann(train_payload, stores_payload, test_payload)
            _render_rossmann(data)
        except Exception as exc:
            st.error(f"Không thể nạp dữ liệu Rossmann: {exc}", icon=":material/error:")
        return

    table_candidates = []
    candidate_errors = []
    for item in uploads:
        if not item.name.lower().endswith((".csv", ".xlsx", ".zip")):
            continue
        try:
            table_candidates.extend(_list_upload_tables(item.getvalue(), item.name))
        except Exception as exc:
            candidate_errors.append(f"{item.name}: {exc}")
    if not table_candidates:
        _page_header("SalesPulse · Dự báo bán hàng", "Tải tệp dữ liệu của bạn để bắt đầu phân tích.")
        if candidate_errors:
            st.error("Không đọc được bảng nào trong file đã tải lên. " + " · ".join(candidate_errors))
        else:
            st.info("Chưa có dữ liệu được nạp. Chọn tệp CSV, XLSX hoặc ZIP ở thanh bên; sau đó ghép các cột ngày và chỉ tiêu cần phân tích.", icon=":material/upload_file:")
        with st.container(border=True):
            st.subheader("Các cột cần có")
            st.markdown("- **Bắt buộc:** một cột ngày và một cột số cần dự báo (doanh thu, sales, quantity, đơn hàng…).\n- **Tùy chọn:** cột số lượng, mã đơn hàng và một cột nhóm như khu vực, cửa hàng hoặc danh mục.\n- **Định dạng:** CSV, XLSX; ZIP, kể cả ZIP lồng ZIP như bộ Walmart. Với bộ nhiều file, chọn bảng lịch sử `train.csv` thay vì `test.csv` hoặc `sample_submission.csv`.")
        return

    try:
        if len(table_candidates) > 1:
            candidate_labels = [item["label"] for item in table_candidates]
            preferred_index = _preferred_table_index(table_candidates)
            with st.sidebar:
                selected_label = st.selectbox("Chọn bảng dữ liệu", candidate_labels,
                                              index=preferred_index, key="upload_table_choice",
                                              help="ZIP có thể chứa nhiều bảng. Chọn bảng lịch sử có cả cột ngày và cột doanh số/số lượng.")
            selected_candidate = table_candidates[candidate_labels.index(selected_label)]
        else:
            selected_candidate = table_candidates[0]
        headers, rows = _read_tabular_upload(selected_candidate["payload"], selected_candidate["filename"])
        headers = _make_unique_headers(headers)
        source_filename = selected_candidate["label"]
        date_aliases = ["date", "order date", "invoice date", "invoicedate", "transaction date", "datetime", "timestamp", "sales date", "datekey", "ds", "ngay", "ngay ban"]
        metric_aliases = ["sales", "revenue", "total sales", "weekly sales", "net sales", "sales amount", "amount paid", "total amount", "turnover", "gmv", "gross income"]
        quantity_aliases = ["quantity", "qty", "units", "item count", "units sold", "quantity sold", "unit_sales", "unit sales"]
        date_default = _suggest_column(headers, date_aliases)
        metric_default = _suggest_column(headers, metric_aliases)
        quantity_default = _suggest_column(headers, quantity_aliases, allow_none=True)
        price_default = _suggest_column(headers, ["unit price", "unitprice", "price per unit", "price"], allow_none=True)
        if metric_default == headers[0] and price_default != "(Bỏ qua)" and quantity_default != "(Bỏ qua)":
            metric_default = DERIVED_REVENUE
        elif metric_default == headers[0]:
            metric_default = _suggest_column(headers, quantity_aliases)
        metric_options = headers + ([DERIVED_REVENUE] if price_default != "(Bỏ qua)" and quantity_default != "(Bỏ qua)" else [])
        if quantity_default == metric_default:
            quantity_default = "(Bỏ qua)"
        group_default = _suggest_column(headers, ["region", "store", "store name", "store id", "country", "city", "category", "product", "family", "department", "item", "branch", "location"], allow_none=True)
        if group_default in (date_default, metric_default, quantity_default):
            group_default = "(Bỏ qua)"
        order_default = _suggest_column(headers, ["order id", "order number", "invoice no", "invoiceno", "transaction id", "receipt id", "id"], allow_none=True)

        st.caption(f"Bảng đang phân tích: `{source_filename}` · {len(rows):,} dòng · {len(headers)} cột")
        with st.expander("Xem trước 5 dòng và tiêu đề cột", expanded=False):
            st.dataframe([dict(zip(headers, row)) for row in rows[:5]], hide_index=True)

        with st.sidebar:
            st.divider()
            st.markdown("**Ghép cột dữ liệu**")
            date_column = st.selectbox("Cột ngày", headers, index=headers.index(date_default), key="upload_date_column")
            date_format_labels = ["Tự nhận diện", "Ngày/tháng/năm", "Tháng/ngày/năm", "Năm-tháng-ngày"]
            date_format_label = st.selectbox("Định dạng ngày", date_format_labels, key="upload_date_format")
            date_patterns = {"Tự nhận diện": None, "Ngày/tháng/năm": "%d/%m/%Y", "Tháng/ngày/năm": "%m/%d/%Y", "Năm-tháng-ngày": "%Y-%m-%d"}
            date_pattern = date_patterns[date_format_label]
            metric_column = st.selectbox("Cột cần dự báo / doanh thu", metric_options,
                                         index=metric_options.index(metric_default) if metric_default in metric_options else 0,
                                         key="upload_metric_column")
            optional = ["(Bỏ qua)"] + headers
            quantity_column = st.selectbox("Cột số lượng · tùy chọn", optional,
                                           index=optional.index(quantity_default) if quantity_default in optional else 0,
                                           key="upload_quantity_column")
            price_column = None
            if metric_column == DERIVED_REVENUE:
                price_column = st.selectbox("Cột đơn giá", headers,
                                            index=headers.index(price_default) if price_default in headers else 0,
                                            key="upload_price_column")
            group_column = st.selectbox("Cột phân nhóm · tùy chọn", optional,
                                        index=optional.index(group_default) if group_default in optional else 0,
                                        key="upload_group_column")
            order_column = st.selectbox("Cột mã đơn hàng · tùy chọn", optional,
                                        index=optional.index(order_default) if order_default in optional else 0,
                                        key="upload_order_column")
            st.divider()
            st.markdown("**Thiết lập dự báo**")
            frequency = st.segmented_control("Gom dữ liệu theo", ["Ngày", "Tuần", "Tháng"], default="Tháng") or "Tháng"
            choices = {"Ngày": [7, 14, 30], "Tuần": [4, 8, 12], "Tháng": [1, 2, 3]}[frequency]
            horizon = st.selectbox("Số kỳ dự báo", choices, index=len(choices) - 1,
                                   format_func=lambda value: f"{value} kỳ")
            model_name = st.selectbox("Mô hình", ["Linear Regression", "Random Forest"])

        with st.spinner("Đang xử lý các cột đã chọn…"):
            data = _aggregate_uploaded_data(headers, rows, date_column, metric_column, quantity_column,
                                            group_column, order_column, source_filename, date_pattern, price_column)
        target_options = ["Revenue"]
        if data["has_quantity"] and quantity_column != metric_column:
            target_options.append("Quantity")
        with st.sidebar:
            target = st.selectbox("Chỉ tiêu dự báo", target_options,
                                  format_func=lambda value: data["metric_label"] if value == "Revenue" else quantity_column,
                                  key="upload_forecast_target")
        required_periods = len(make_series(data, frequency, target)[0])
        if required_periods < 15:
            st.warning(f"Sau khi tổng hợp còn {required_periods} kỳ. Cần ít nhất 15 kỳ theo mức gộp đã chọn; thử chọn Ngày hoặc Tuần nếu dữ liệu có đủ ngày.")
            return
        _render_uploaded_pages(data, frequency, horizon, model_name, target)
    except Exception as exc:
        st.warning("Chưa thể phân tích bảng đã chọn.", icon=":material/warning:")
        st.write("Chọn bảng lịch sử (thường là `train.csv`), rồi kiểm tra cột ngày, cột số cần dự báo và định dạng ngày. Dữ liệu cần có ít nhất 15 kỳ sau khi tổng hợp.")
        st.caption(f"Chi tiết: {exc}")


if __name__ == "__main__":
    main()
