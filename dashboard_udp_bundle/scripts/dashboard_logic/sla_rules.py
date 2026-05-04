"""SLA rules used by dashboard metrics."""

SLA_VIOLATED_VALUE = 'нарушен'


def is_sla_violated(value):
    return value == SLA_VIOLATED_VALUE
