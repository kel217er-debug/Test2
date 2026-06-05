"""Excel column mapping and row-level filtering rules for 'Заявки+обращения' dataset.

Отдельно от `excel_filters.py`, чтобы не менять "МУЗ"-проект.
"""

# Каналы (по фактическим значениям в объединённом файле).
# Здесь используем максимально широкий набор, чтобы метрики строились по всем строкам.
OUR_CHANNELS = {"ЦК РП", "ГАП", "ГПО", "3К", "ГПХ"}

# Для объединённого файла направления как в МУЗ ("НОД/НПГС") не применимы.
DIR_MAP = {}

EXCL_SERVICES = {"другие услуги", "другая услуга", "заказ оборудования"}
EMPLOYEE_TAB_EXCL_PATTERNS = ("заказ оборудования", "другие услуги")
APRIL_2026_FEDERAL_CHANNEL_EXCL_MRFS = {"РД Центр", "РД ЮГ", "РД Волга"}
CONNECTED_STATUS = "Услуга подключена"
INET_LOWER = "интернет"

# Индексы колонок (оставляем как в базовом проекте — синтетический экспорт строится под них).
COL = {
    "mrf": 0,
    "region": 1,
    "reg_dt": 10,
    "channel": 12,
    "service": 19,
    "equip_price": 24,
    "inn": 31,
    "segment": 33,
    "exec_name": 50,
    "exec_podr": 51,
    "accept_dt": 52,
    "primary_name": 53,
    "primary_podr": 54,
    "sla_cont": 66,
    "sla_acc": 67,
    "transfer_dt": 70,
    "connection_result": 120,
    "connector_name": 121,
    "connected_services_cnt": 124,
    "install_amount": 125,
    "monthly_amount": 126,
    "final_dt": 127,
    "final_status": 128,
    "final_reason": 129,
    "current_status": 133,
    "current_exec": 134,
    "current_podr": 135,
    "hrs_since_reg": 136,
    "hrs_in_status": 137,
}


def is_our_channel(podr: str) -> bool:
    return (podr or "") in OUR_CHANNELS


def direction_for_channel(podr: str) -> str:
    return DIR_MAP.get(podr, "")


def is_base_service(service_lower: str) -> bool:
    return (service_lower or "") not in EXCL_SERVICES


def normalize_service_name(service_value: str) -> str:
    return " ".join(str(service_value or "").replace("\xa0", " ").strip().lower().split())


def is_employee_tab_excluded_service(service_value: str) -> bool:
    service_norm = normalize_service_name(service_value)
    return any(pattern in service_norm for pattern in EMPLOYEE_TAB_EXCL_PATTERNS)


def is_employee_tab_excluded_april_2026_federal_channel(reg_dt, channel_value: str, employee_mrf: str) -> bool:
    if reg_dt is None or getattr(reg_dt, "year", None) != 2026 or getattr(reg_dt, "month", None) != 4:
        return False
    channel_norm = " ".join(str(channel_value or "").replace("\xa0", " ").strip().split())
    employee_mrf_norm = " ".join(str(employee_mrf or "").replace("\xa0", " ").strip().split())
    return channel_norm == "Партнеры федеральные" and employee_mrf_norm in APRIL_2026_FEDERAL_CHANNEL_EXCL_MRFS


def is_transferred(current_podr: str) -> bool:
    # Для объединённого источника считаем передачей уход из нашего списка.
    return (current_podr or "") not in OUR_CHANNELS


def is_connected(final_status: str) -> bool:
    return (final_status or "") == CONNECTED_STATUS


def is_internet_service(service_lower: str) -> bool:
    s = (service_lower or "").strip()
    return s == INET_LOWER or INET_LOWER in s
