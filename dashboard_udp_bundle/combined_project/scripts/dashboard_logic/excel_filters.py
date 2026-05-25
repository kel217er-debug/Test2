"""Excel column mapping and row-level filtering rules."""

OUR_CHANNELS = {'НОД', 'Партнеры НОД', 'НПГС'}
DIR_MAP = {'НОД': 'НИР', 'Партнеры НОД': 'НПП', 'НПГС': 'НСП'}
EXCL_SERVICES = {'другие услуги', 'другая услуга', 'заказ оборудования'}
CONNECTED_STATUS = 'Клиент подключен'
INET_LOWER = 'интернет'

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


def is_transferred(current_podr):
    return current_podr not in OUR_CHANNELS


def is_connected(final_status):
    return final_status == CONNECTED_STATUS


def is_internet_service(service_lower):
    return service_lower == INET_LOWER
