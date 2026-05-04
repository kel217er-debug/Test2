# Ежедневный дашборд УДП

Проект собирает автономный HTML-дашборд по Excel-выгрузке из МУЗ. Данные обрабатываются Python-скриптом, сохраняются в JSON, затем встраиваются в HTML-шаблон вместе с локальной копией Chart.js. Итоговый файл можно открыть в браузере без интернета и без запуска сервера.

## Быстрый запуск обновления

Из папки `dashboard_udp_bundle` выполните:

```powershell
python scripts/update_daily_dashboard.py
```

После успешного запуска обновляются:

- `data/processed/daily_dashboard_data.json` - рассчитанные данные для дашборда.
- `dist/udp_daily_dashboard.html` - готовый HTML-файл для просмотра.

Итоговый дашборд открывается двойным кликом по `dist/udp_daily_dashboard.html`.

## Структура проекта

```text
dashboard_udp_bundle/
├── config/
│   └── employee_hierarchy.json
├── data/
│   ├── source/
│   │   └── muz_export.xlsx
│   └── processed/
│       └── daily_dashboard_data.json
├── dist/
│   ├── udp_daily_dashboard.html
│   └── udp_daily_dashboard_legacy_2026-04-16.html
├── scripts/
│   ├── update_daily_dashboard.py
│   ├── prepare_dashboard_data.py
│   ├── build_offline_dashboard.py
│   └── dashboard_logic/
│       ├── excel_filters.py
│       ├── sla_rules.py
│       ├── periods.py
│       ├── aggregations.py
│       └── metrics.py
├── templates/
│   └── daily_dashboard_template.html
├── vendor/
│   └── chartjs.umd.js
├── __pycache__/
└── README.md
```

## За что отвечает каждая папка

| Папка | Назначение |
|---|---|
| `config/` | Ручные справочники и настройки. Сейчас здесь хранится иерархия сотрудников. |
| `data/source/` | Исходные входные данные. Сюда кладется свежая Excel-выгрузка из МУЗ. |
| `data/processed/` | Промежуточный результат обработки. JSON для встраивания в HTML. |
| `dist/` | Готовые HTML-файлы, которые можно открывать или передавать пользователям. |
| `scripts/` | Python-скрипты обновления, обработки данных и сборки HTML. |
| `scripts/dashboard_logic/` | Отдельные модули правил ETL: фильтры Excel, SLA, периоды, агрегации и расчет метрик. |
| `templates/` | HTML-шаблон интерфейса дашборда: разметка, стили и JavaScript-логика. |
| `vendor/` | Локальные сторонние библиотеки. Нужны для автономной работы без интернета. |
| `__pycache__/` | Служебный кеш Python. Создается автоматически, вручную не редактируется. |

## Что делает каждый файл

| Файл | Действие |
|---|---|
| `scripts/update_daily_dashboard.py` | Главная команда обновления. Проверяет наличие обязательных файлов, ищет Excel-файл в `data/source/`, запускает подготовку JSON, затем сборку HTML. Если подготовка данных завершилась ошибкой, сборка HTML не запускается. |
| `scripts/prepare_dashboard_data.py` | Главный ETL-скрипт-оркестратор. Читает Excel, проходит по строкам, вызывает отдельные модули из `scripts/dashboard_logic/` и сохраняет результат в `data/processed/daily_dashboard_data.json`. |
| `scripts/build_offline_dashboard.py` | Собирает автономный HTML. Берет JSON из `data/processed/`, шаблон из `templates/`, локальный Chart.js из `vendor/`, подставляет данные в плейсхолдеры и записывает `dist/udp_daily_dashboard.html`. |
| `scripts/dashboard_logic/excel_filters.py` | Хранит индексы колонок Excel (`COL`) и правила фильтрации строк: нужные каналы, направления, исключаемые услуги, статус подключения, определение передач. |
| `scripts/dashboard_logic/sla_rules.py` | Хранит правила SLA. Сейчас определяет, какое значение в Excel считается нарушением SLA. |
| `scripts/dashboard_logic/periods.py` | Считает недельные и месячные ключи, границы периодов, подписи месяцев, список периодов и timeline для дашборда. |
| `scripts/dashboard_logic/aggregations.py` | Создает пустые структуры агрегаций для регистраций, закрытий, открытых заявок и сериализует недельные/месячные срезы. |
| `scripts/dashboard_logic/metrics.py` | Считает и преобразует метрики: заявки, передачи, SLA, подключения, НД, оборудование, допродажи, фрод-индикатор, открытые и застрявшие заявки. |
| `config/employee_hierarchy.json` | Справочник сотрудников: ФИО, тимлид, руководитель, направление, МРФ, активность. Его нужно обновлять при изменении состава команд или структуры подчинения. |
| `data/source/muz_export.xlsx` | Исходная Excel-выгрузка. При обновлении данных этот файл заменяется свежей выгрузкой. Скрипты ищут `.xlsx`, в названии которого есть `muz` или `муз`. |
| `data/processed/daily_dashboard_data.json` | Машиночитаемый результат ETL. Используется HTML-сборщиком. Обычно вручную не редактируется, потому что перезаписывается при обновлении. |
| `templates/daily_dashboard_template.html` | Основной HTML-шаблон дашборда. Содержит вкладки, фильтры, таблицы, графики, стили и клиентскую JavaScript-логику. Внутри есть плейсхолдеры `/*__DATA__*/`, `/*__CHARTJS__*/`, `/*__BUILD_TIME__*/`. |
| `vendor/chartjs.umd.js` | Локальная копия Chart.js. Благодаря ей графики работают без подключения к интернету. |
| `dist/udp_daily_dashboard.html` | Актуальный готовый дашборд. Это главный файл, который открывается после обновления. |
| `dist/udp_daily_dashboard_legacy_2026-04-16.html` | Архивная старая версия дашборда. В текущей сборке не используется, но может быть полезна для сравнения или восстановления. |
| `README.md` | Описание проекта, структура, назначение файлов и инструкция по обновлению. |

## Как обновлять проект

1. Откройте папку проекта:

```powershell
cd "c:\Users\weren\OneDrive\Рабочий стол\Рабочие скрипты python\Test2\dashboard_udp_bundle"
```

2. Положите свежую Excel-выгрузку в папку `data/source/`.

3. Убедитесь, что файл имеет расширение `.xlsx`, а в названии есть `МУЗ`.

4. Если структура сотрудников изменилась, обновите `config/employee_hierarchy.json`.

5. Запустите полное обновление:

```powershell
python scripts/update_daily_dashboard.py
```

6. Проверьте, что в консоли нет строки `FAILED`.

7. Откройте результат:

```text
dist/udp_daily_dashboard.html
```

## Ручное обновление по шагам

Обычно достаточно запускать `scripts/update_daily_dashboard.py`. Ручные команды нужны, если требуется выполнить только часть процесса.

Пересчитать данные из Excel в JSON:

```powershell
python scripts/prepare_dashboard_data.py
```

Пересобрать HTML из уже готового JSON:

```powershell
python scripts/build_offline_dashboard.py
```

Правильный порядок при полном ручном обновлении:

1. Сначала `python scripts/prepare_dashboard_data.py`.
2. Затем `python scripts/build_offline_dashboard.py`.

## Когда какой файл редактировать

| Задача | Где править |
|---|---|
| Обновить входные данные | Заменить Excel в `data/source/`. |
| Добавить сотрудника или изменить тимлида, руководителя, направление, МРФ | `config/employee_hierarchy.json`. |
| Изменить расчет метрик | `scripts/dashboard_logic/metrics.py`. |
| Изменить фильтры строк Excel или индексы колонок | `scripts/dashboard_logic/excel_filters.py`. |
| Изменить правила SLA | `scripts/dashboard_logic/sla_rules.py`. |
| Изменить периоды, подписи недель/месяцев или timeline | `scripts/dashboard_logic/periods.py`. |
| Изменить структуры агрегаций и сериализацию срезов | `scripts/dashboard_logic/aggregations.py`. |
| Изменить общий проход ETL по строкам Excel | `scripts/prepare_dashboard_data.py`. |
| Изменить внешний вид, вкладки, таблицы, фильтры, графики или поведение интерфейса | `templates/daily_dashboard_template.html`. |
| Изменить способ сборки HTML или имя итогового файла | `scripts/build_offline_dashboard.py`. |
| Изменить общий порядок обновления и проверки входных файлов | `scripts/update_daily_dashboard.py`. |
| Обновить библиотеку графиков | `vendor/chartjs.umd.js`. |

## Логика обработки данных

Основной источник данных - Excel-файл из `data/source/`. Скрипт `prepare_dashboard_data.py` читает лист `Sheet1` через пакет `python-calamine`.

Колонки Excel берутся по фиксированным индексам из словаря `COL` внутри `scripts/dashboard_logic/excel_filters.py`. Если в выгрузке изменится порядок колонок, нужно обновить индексы в этом словаре.

Скрипт считает два независимых набора метрик:

- `registered` - входящий поток по дате регистрации заявки: количество заявок, SLA, передачи, нагрузка, открытые и застрявшие заявки, когортная конверсия.
- `closed` - результат по дате перевода в итоговый статус: закрытия, подключения, НД, услуги, оборудование, Close%, допродажи и фрод-индикатор.

Периоды делятся на:

- закрытые месяцы - показываются как месячные итоги;
- текущий открытый месяц - показывается по неделям.

## Справочник сотрудников

`config/employee_hierarchy.json` связывает сотрудника с командной структурой:

- `name` - ФИО сотрудника;
- `teamlead` - тимлид;
- `director` - руководитель;
- `direction` - направление;
- `mrf` - МРФ;
- `is_active` - активен ли сотрудник.

Если в новой Excel-выгрузке появится сотрудник, которого нет в справочнике, он попадет в данные, но может отображаться без привязки к команде. Чтобы фильтры и командные таблицы работали корректно, добавьте такого сотрудника в `employee_hierarchy.json`.

## Требования

- Python 3.10 или новее.
- Пакет `python-calamine`.

Установка зависимости:

```powershell
pip install python-calamine
```

## Проверка после обновления

После запуска `python scripts/update_daily_dashboard.py` проверьте:

1. В консоли оба шага завершились статусом `OK`.
2. Нет строки `FAILED`.
3. Обновился файл `data/processed/daily_dashboard_data.json`.
4. Обновился файл `dist/udp_daily_dashboard.html`.
5. `dist/udp_daily_dashboard.html` открывается в браузере и показывает свежую дату сборки.

## Частые проблемы

| Симптом | Что проверить |
|---|---|
| `missing required files` | На месте ли `config/employee_hierarchy.json`, шаблон, Chart.js и оба рабочих скрипта. |
| Не найден Excel | В `data/source/` должен лежать `.xlsx`-файл, в названии которого есть `muz` или `муз`. |
| Ошибка импорта `python_calamine` | Установите зависимость командой `pip install python-calamine`. |
| HTML собрался, но данные старые | Сначала запустите `prepare_dashboard_data.py` или полный `update_daily_dashboard.py`. |
| Графики не работают офлайн | Проверьте наличие `vendor/chartjs.umd.js`. |
