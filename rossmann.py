"""Rossmann CSV loading and NumPy-free forecasting for the SalesPulse demo."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from io import StringIO
import csv
import math
import random


FEATURE_NAMES = [
    "Doanh thu hôm trước", "Cùng thứ tuần trước", "Cùng thứ 2 tuần trước", "Cùng thứ 4 tuần trước",
    "Trung bình 7 ngày", "Trung bình 28 ngày", "Xu hướng", "Chu kỳ thứ (sin)", "Chu kỳ thứ (cos)",
    "Chu kỳ tháng (sin)", "Chu kỳ tháng (cos)", "Cửa hàng mở", "Khuyến mãi", "Ngày lễ quốc gia", "Nghỉ học",
]


def _rows(payload: bytes):
    text = payload.decode("utf-8-sig", errors="replace")
    return csv.DictReader(StringIO(text))


def _date(value: str) -> date | None:
    value = str(value or "").strip()
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(value, pattern).date()
        except ValueError:
            pass
    return None


def _float(value, default=0.0) -> float:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return default


def _flag(value, default=0) -> int:
    text = str(value or "").strip().lower()
    if text in {"1", "true", "yes", "y"}:
        return 1
    if text in {"0", "false", "no", "n", ""}:
        return default if text == "" else 0
    # Rossmann StateHoliday is 0 for none, and a/b/c for holiday types.
    return {"a": 1, "b": 2, "c": 3}.get(text, default)


def is_rossmann_csv(payload: bytes) -> bool:
    try:
        reader = _rows(payload)
        columns = {str(item).strip().lower() for item in reader.fieldnames or []}
    except (StopIteration, AttributeError, TypeError):
        return False
    # Require Rossmann-specific context fields so generic store/date/sales tables
    # continue through the user's flexible column-mapping flow.
    return {"store", "date", "sales", "dayofweek", "open", "promo", "stateholiday", "schoolholiday"}.issubset(columns)


def load_rossmann(train_payload: bytes, stores_payload: bytes | None = None,
                  test_payload: bytes | None = None) -> dict:
    """Read Kaggle train/store/test files with the Python standard library only."""
    first_day, last_day = None, None
    store_ids = set()
    for row in _rows(train_payload):
        day = _date(row.get("Date"))
        if day is None:
            continue
        first_day = day if first_day is None or day < first_day else first_day
        last_day = day if last_day is None or day > last_day else last_day
        store_ids.add(str(row.get("Store", "")).strip())
    if first_day is None or not store_ids:
        raise ValueError("train.csv cần các cột Store, Date, Sales và các ngày hợp lệ.")

    dates = [first_day + timedelta(days=offset) for offset in range((last_day - first_day).days + 1)]
    store_data = {
        store_id: {
            "sales": [0.0] * len(dates), "open": [0] * len(dates), "promo": [0] * len(dates),
            "state_holiday": [0] * len(dates), "school_holiday": [0] * len(dates),
            "customers": [0.0] * len(dates),
        }
        for store_id in store_ids
    }
    summary = {store_id: {"sales": 0.0, "customers": 0.0, "promo_sales": 0.0,
                          "promo_days": 0, "open_days": 0, "records": 0}
               for store_id in store_ids}
    for row in _rows(train_payload):
        store_id = str(row.get("Store", "")).strip()
        day = _date(row.get("Date"))
        if not day or store_id not in store_data:
            continue
        index = (day - first_day).days
        sales = max(0.0, _float(row.get("Sales")))
        opened = _flag(row.get("Open"), 1)
        promo = _flag(row.get("Promo"))
        data = store_data[store_id]
        data["sales"][index] = sales
        data["open"][index] = opened
        data["promo"][index] = promo
        data["state_holiday"][index] = _flag(row.get("StateHoliday"))
        data["school_holiday"][index] = _flag(row.get("SchoolHoliday"))
        data["customers"][index] = max(0.0, _float(row.get("Customers")))
        info = summary[store_id]
        info["sales"] += sales
        info["customers"] += data["customers"][index]
        info["promo_sales"] += sales if promo and opened else 0.0
        info["promo_days"] += 1 if promo and opened else 0
        info["open_days"] += 1 if opened else 0
        info["records"] += 1

    metadata = {}
    if stores_payload:
        for row in _rows(stores_payload):
            store_id = str(row.get("Store", "")).strip()
            if store_id:
                metadata[store_id] = {key: str(value or "").strip() for key, value in row.items()}

    future_context = {}
    test_dates = set()
    if test_payload:
        for row in _rows(test_payload):
            store_id = str(row.get("Store", "")).strip()
            day = _date(row.get("Date"))
            if not store_id or day is None:
                continue
            future_context[(store_id, day)] = {
                "open": _flag(row.get("Open"), 1),
                "promo": _flag(row.get("Promo")),
                "state_holiday": _flag(row.get("StateHoliday")),
                "school_holiday": _flag(row.get("SchoolHoliday")),
            }
            test_dates.add(day)
    return {"kind": "rossmann", "dates": dates, "stores": store_data,
            "summary": summary, "metadata": metadata, "future_context": future_context,
            "test_dates": sorted(test_dates), "rows": sum(item["records"] for item in summary.values())}


def _features(history: list[float], day: date, step: int, context: dict) -> list[float]:
    result = [history[-1], history[-7], history[-14], history[-28],
              sum(history[-7:]) / 7, sum(history[-28:]) / 28]
    result.extend([step / 100.0, math.sin(day.weekday() * 2 * math.pi / 7),
                   math.cos(day.weekday() * 2 * math.pi / 7),
                   math.sin(day.month * 2 * math.pi / 12), math.cos(day.month * 2 * math.pi / 12),
                   float(context.get("open", 1)), float(context.get("promo", 0)),
                   float(context.get("state_holiday", 0)), float(context.get("school_holiday", 0))])
    return result


def _solve(matrix: list[list[float]], vector: list[float]) -> list[float]:
    size = len(vector)
    augmented = [matrix[i][:] + [vector[i]] for i in range(size)]
    for column in range(size):
        pivot = max(range(column, size), key=lambda row: abs(augmented[row][column]))
        augmented[column], augmented[pivot] = augmented[pivot], augmented[column]
        if abs(augmented[column][column]) < 1e-12:
            augmented[column][column] = 1e-8
        divisor = augmented[column][column]
        for j in range(column, size + 1):
            augmented[column][j] /= divisor
        for row in range(size):
            if row == column:
                continue
            factor = augmented[row][column]
            for j in range(column, size + 1):
                augmented[row][j] -= factor * augmented[column][j]
    return [augmented[i][-1] for i in range(size)]


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
    if depth >= 4 or len(y) < 18:
        return ("leaf", mean)
    width = len(x[0])
    features = rng.sample(range(width), max(2, int(math.sqrt(width))))
    best = None
    min_leaf = 7
    for feature in features:
        ordered = sorted((row[feature], target) for row, target in zip(x, y))
        count = len(ordered)
        stride = max(1, count // 10)
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
    lx, ly, rx, ry = [], [], [], []
    for row, target in zip(x, y):
        (lx if row[feature] <= threshold else rx).append(row)
        (ly if row[feature] <= threshold else ry).append(target)
    if not ly or not ry:
        return ("leaf", mean)
    return ("node", feature, threshold, _fit_tree(lx, ly, rng, depth + 1), _fit_tree(rx, ry, rng, depth + 1))


def _tree_predict(tree, row: list[float]) -> float:
    while tree[0] == "node":
        tree = tree[3] if row[tree[1]] <= tree[2] else tree[4]
    return tree[1]


def _fit_forest(x: list[list[float]], y: list[float]):
    rng = random.Random(42)
    result = []
    for _ in range(18):
        sample = [rng.randrange(len(y)) for _ in range(len(y))]
        result.append(_fit_tree([x[i] for i in sample], [y[i] for i in sample], rng))
    return result


def _model_importance(model, name: str) -> list[dict]:
    if name == "Linear Regression":
        values = [abs(weight) for weight in model["weights"][1:]]
        method = "Trị tuyệt đối hệ số chuẩn hóa"
    else:
        values = [0.0] * len(FEATURE_NAMES)
        def visit(tree):
            if tree[0] == "node":
                values[tree[1]] += 1.0
                visit(tree[3])
                visit(tree[4])
        for tree in model:
            visit(tree)
        method = "Tỷ trọng số nút chia trong rừng"
    total = sum(values) or 1.0
    return [{"feature": feature, "importance": value / total * 100, "method": method}
            for feature, value in zip(FEATURE_NAMES, values)]


def _predict_model(model, row: list[float], name: str) -> float:
    if name == "Linear Regression":
        return _predict_linear(model, row)
    return sum(_tree_predict(tree, row) for tree in model) / len(model)


def _rollout(model, name: str, history: list[float], dates: list[date], contexts: list[dict]):
    history = history[:]
    predicted = []
    for day, context in zip(dates, contexts):
        row = _features(history, day, len(history), context)
        value = 0.0 if not context.get("open", 1) else max(0.0, _predict_model(model, row, name))
        predicted.append(value)
        history.append(value)
    return predicted


def _baseline(history: list[float], contexts: list[dict]) -> list[float]:
    result = []
    for index, context in enumerate(contexts):
        seasonal = history[-7 + (index % 7)]
        result.append(seasonal if context.get("open", 1) else 0.0)
    return result


def _score(actual: list[float], predicted: list[float]) -> dict:
    errors = [abs(a - p) for a, p in zip(actual, predicted)]
    nonzero = [(a, p) for a, p in zip(actual, predicted) if a > 0]
    rmspe = math.sqrt(sum(((a - p) / a) ** 2 for a, p in nonzero) / len(nonzero)) * 100 if nonzero else 0.0
    denominator = sum(abs(value) for value in actual)
    wape = sum(errors) / denominator * 100 if denominator else 0.0
    return {"mae": sum(errors) / len(errors),
            "rmse": math.sqrt(sum((a - p) ** 2 for a, p in zip(actual, predicted)) / len(errors)),
            "wape": wape, "rmspe": rmspe}


def evaluate_and_forecast(dates: list[date], values: list[float], contexts: list[dict],
                          horizon: int, future_dates: list[date], future_contexts: list[dict]) -> dict:
    first = 28
    if len(values) <= first + 20:
        raise ValueError("Cần ít nhất khoảng 50 ngày lịch sử cho huấn luyện và kiểm tra.")
    holdout = min(42, max(14, horizon), max(14, len(values) // 5))
    split = len(values) - holdout
    x, y = [], []
    for index in range(first, len(values)):
        x.append(_features(values[:index], dates[index], index, contexts[index]))
        y.append(values[index])
    train_count = split - first
    x_train, y_train = x[:train_count], y[:train_count]
    if len(x_train) < 10:
        raise ValueError("Không đủ ngày lịch sử trước kỳ kiểm tra.")

    validation_dates = dates[split:]
    validation_contexts = contexts[split:]
    actual = values[split:]
    results = {"Seasonal baseline": {
        "validation": _baseline(values[:split], validation_contexts),
        "forecast": _baseline(values, future_contexts),
    }}
    for name in ("Linear Regression", "Random Forest"):
        if name == "Linear Regression":
            validation_model = _fit_linear(x_train, y_train)
            final_model = _fit_linear(x, y)
        else:
            validation_model = _fit_forest(x_train, y_train)
            final_model = _fit_forest(x, y)
        results[name] = {
            "validation": _rollout(validation_model, name, values[:split], validation_dates, validation_contexts),
            "forecast": _rollout(final_model, name, values, future_dates, future_contexts),
            "importance": _model_importance(final_model, name),
        }
    for result in results.values():
        result["metrics"] = _score(actual, result["validation"])
    return {"validation_dates": validation_dates, "actual": actual, "forecast_dates": future_dates,
            "results": results, "holdout": holdout}


def store_contexts(store: dict) -> list[dict]:
    return [{"open": store["open"][i], "promo": store["promo"][i],
             "state_holiday": store["state_holiday"][i], "school_holiday": store["school_holiday"][i]}
            for i in range(len(store["sales"]))]


def promotion_summary(store: dict) -> dict:
    promo_days = [store["sales"][i] for i in range(len(store["sales"]))
                  if store["open"][i] and store["promo"][i]]
    regular_days = [store["sales"][i] for i in range(len(store["sales"]))
                    if store["open"][i] and not store["promo"][i]]
    promo_avg = sum(promo_days) / len(promo_days) if promo_days else 0.0
    regular_avg = sum(regular_days) / len(regular_days) if regular_days else 0.0
    uplift = (promo_avg / regular_avg - 1) * 100 if regular_avg else 0.0
    return {"promo_avg": promo_avg, "regular_avg": regular_avg,
            "uplift": uplift, "promo_days": len(promo_days), "regular_days": len(regular_days)}


def weekday_summary(store: dict, dates: list[date]) -> list[dict]:
    labels = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
    values = [[store["sales"][i] for i, day in enumerate(dates) if day.weekday() == weekday and store["open"][i]]
              for weekday in range(7)]
    return [{"day": labels[i], "average": sum(group) / len(group) if group else 0.0, "count": len(group)}
            for i, group in enumerate(values)]
