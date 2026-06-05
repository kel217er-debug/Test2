"""Excel column mapping and row-level filtering rules."""

OUR_CHANNELS = {'НОД', 'Партнеры НОД', 'НПГС'}
DIR_MAP = {'НОД': 'НИР', 'Партнеры НОД': 'НПП', 'НПГС': 'НСП'}
EXCL_SERVICES = {'другие услуги', 'другая услуга', 'заказ оборудования'}
EMPLOYEE_TAB_EXCL_PATTERNS = ('заказ оборудования', 'другие услуги')
APRIL_2026_FEDERAL_CHANNEL_EXCL_MRFS = {'РД Центр', 'РД ЮГ', 'РД Волга'}
CONNECTED_STATUS = 'Клиент подключен'
INET_LOWER = 'интернет'
MUZ_RD_CENTER_MRF = 'Центр'
MUZ_RD_CENTER_EXCL_CHANNEL = 'НОД'
MUZ_RD_CENTER_EXCL_SERVICES = {'виртуальная атс', 'номер 8800'}

COL = {
    'mrf': 0,
    'region': 1,
    'reg_dt': 10,
    'channel': 12,
    'service': 19,
    'equip_price': 24,
    'inn': 31,
    'segment': 33,
    'exec_name': 50,
    'exec_podr': 51,
    'accept_dt': 52,
    'primary_name': 53,
    'primary_podr': 54,
    'sla_cont': 66,
    'sla_acc': 67,
    'transfer_dt': 70,
    'connection_result': 120,
    'connector_name': 121,
    'connected_services_cnt': 124,
    'install_amount': 125,
    'monthly_amount': 126,
    'final_dt': 127,
    'final_status': 128,
    'final_reason': 129,
    'current_status': 133,
    'current_exec': 134,
    'current_podr': 135,
    'hrs_since_reg': 136,
    'hrs_in_status': 137,
}


def is_our_channel(podr):
    return podr in OUR_CHANNELS


def direction_for_channel(podr):
    return DIR_MAP.get(podr, '')


def is_base_service(service_lower):
    return service_lower not in EXCL_SERVICES


def normalize_service_name(service_value):
    return ' '.join(str(service_value or '').replace('\xa0', ' ').strip().lower().split())


def normalize_text_value(value):
    return ' '.join(str(value or '').replace('\xa0', ' ').strip().split())


def is_employee_tab_excluded_service(service_value):
    service_norm = normalize_service_name(service_value)
    return any(pattern in service_norm for pattern in EMPLOYEE_TAB_EXCL_PATTERNS)


def is_excluded_muz_rd_center_nod_service(source_mrf, processing_channel, service_value):
    mrf_norm = normalize_text_value(source_mrf)
    channel_norm = normalize_text_value(processing_channel)
    service_norm = normalize_service_name(service_value)
    return (
        mrf_norm == MUZ_RD_CENTER_MRF
        and channel_norm == MUZ_RD_CENTER_EXCL_CHANNEL
        and service_norm in MUZ_RD_CENTER_EXCL_SERVICES
    )


def is_employee_tab_excluded_april_2026_federal_channel(reg_dt, channel_value, employee_mrf):
    if reg_dt is None or getattr(reg_dt, 'year', None) != 2026 or getattr(reg_dt, 'month', None) != 4:
        return False
    channel_norm = normalize_text_value(channel_value)
    employee_mrf_norm = normalize_text_value(employee_mrf)
    return channel_norm == 'Партнеры федеральные' and employee_mrf_norm in APRIL_2026_FEDERAL_CHANNEL_EXCL_MRFS


def is_transferred(current_podr):
    return current_podr not in OUR_CHANNELS


def is_connected(final_status):
    return final_status == CONNECTED_STATUS


def is_internet_service(service_lower):
    return service_lower == INET_LOWER
