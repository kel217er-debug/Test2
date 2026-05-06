"""Хелперы для метрики повторных обращений (по комментариям + группировка по телефону)."""

import re


# Подстроки намеренно "нестрогие" (стемы), чтобы матчить разные формы слов.
REPEAT_KEYWORDS = (
    'повтор',
    'срочно',
    'эскалац',
    'ждёт звонка',
    'ждет звонка',
    'ждут звонка',
    ' ос',
    'связаться',
)


_NON_DIGITS_RE = re.compile(r'\D+')


def normalize_phone(v):
    """Нормализует телефон: оставляет только цифры (если цифр нет — пустая строка)."""
    if v is None:
        return ''
    s = str(v).strip()
    if not s:
        return ''
    return _NON_DIGITS_RE.sub('', s)


def is_repeat_by_comment(comment_value):
    """True, если комментарий содержит одно из ключевых слов (без учёта регистра)."""
    if comment_value is None:
        return False
    s = str(comment_value).strip().lower()
    if not s:
        return False
    return any(k in s for k in REPEAT_KEYWORDS)
