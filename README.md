# Secondary Sales Forecast 2026

Прогноз вторичных продаж (пролонгаций) на 2026 год.  
Пакетная revenue-модель с калибровкой из исторических данных.

---

## ⚠️ Важные правила данных

> **НИКОГДА не использовать поле `Package`** — это заглушка, у многих записей `Package=999`.  
> Всегда использовать **`credits_recieved_package`** (реальный размер пакета в уроках).  
> Фильтр: `credits_recieved_package.abs().between(1, 500)`

---

## Файлы

| Файл | Назначение |
|------|-----------|
| `build_forecast_inputs.py` | Шаг 1: читает CSV → калибрует → создаёт `forecast_inputs_2026.xlsx` |
| `forecast_inputs_2026.xlsx` | Все параметры модели (редактируется вручную) |
| `forecast_2026_v3.py` | Шаг 2: читает Excel → строит прогноз → `forecast_2026_v3.xlsx` |

---

## Быстрый старт

### 1. Установить зависимости
```bash
pip install -r requirements.txt
```

### 2. Указать путь к данным

Открой оба `.py` файла и замени `DATA_PATH` в начале файла на путь к своему CSV:

```python
# В build_forecast_inputs.py и forecast_2026_v3.py
DATA_PATH = r"C:\путь\к\prolongation_test_YYYY-MM-DD.csv"
```

### 3. Сгенерировать inputs (первый раз или при обновлении данных)
```bash
python build_forecast_inputs.py
```
Создаёт `forecast_inputs_2026.xlsx` со всеми откалиброванными параметрами.

### 4. (Опционально) Изменить параметры в Excel

Открой `forecast_inputs_2026.xlsx`, перейди на нужный лист, измени жёлтые ячейки:

| Лист | Что менять |
|------|-----------|
| **`settings`** | Период прогноза, data_cutoff, AOV-коэффициент |
| **`plan_counts`** | Кол-во новых студентов по месяцам |
| **`package_dist`** | Распределение по пакетам (8/16/32 уроков) |
| **`renewal_price`** | Цена за урок по пакету |
| **`retention_by_renewal`** | Retention по номеру оплаты |
| **`shares`** | Доли Present/Earlier/Reanim/Upgrades |
| **`ext_curve`** | Лаговая кривая для Jun-Dec |
| **`rates`** | Flat prolongation rate (fallback) |

### 5. Запустить прогноз
```bash
python forecast_2026_v3.py
```
Создаёт `forecast_2026_v3.xlsx` с 4 листами:
- **Monthly Summary** — итог по месяцам с факт. данными
- **Dim Breakdown** — разбивка по dim (Base/MT × Private/Premium/Group)
- **Calibration** — параметры калибровки
- **Wide Format** — широкий формат для импорта в финмодель

---

## Ключевые параметры (лист `settings`)

### Изменить период прогноза
```
forecast_start = 2026-06     # первый месяц
forecast_end   = 2027-05     # последний (пример: год вперёд)
```

### Скорректировать AOV
```
aov_adjustment_factor = 1.10   # ×1.10 = +10% к выручке
```
Дефолт 1.10 откалиброван по бэктесту: январь 2026 → ошибка -0.2%.

---

## Архитектура пула (3 режима)

| Период | Режим | Формула |
|--------|-------|---------|
| T ≤ Apr 2026 | `data` | `pool_by_payment_no × retention[n]` |
| May 2026 | `blend` | то же × (31/28) для 3 недостающих дней |
| Jun 2026+ | `ext` | `Σ cohort_size × ext_curve[lag]` |

---

## Расчёт выручки

```
revenue = Σ_seg Σ_pkg  total_sec × share[seg] × pkg_dist[dim,pkg]
                       × pkg_size × price_per_lesson[dim,pkg]
                       × price_growth_factor(month)
                       × aov_adjustment_factor
```

`price_growth_factor` = `(1 + growth_pct)^quarters_since_Q3_2025`  
По умолчанию: 3% в квартал.
