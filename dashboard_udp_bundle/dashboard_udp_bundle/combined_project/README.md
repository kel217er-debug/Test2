# Ежедневный дашборд УДП

Единый проект собирает автономный HTML-дашборд из двух источников:

- `data/source/muz.xlsx` — основной источник для всех РД, кроме `РД Центр`.
- `lm/Заявки+обращения_объединено.xlsx` — источник только для `РД Центр`.

Промежуточный расчет для `РД Центр` сохраняется в `lm/rd_center_daily_dashboard_data.json`.

## Запуск

```powershell
python scripts/update_daily_dashboard_combined.py
```

После успешного запуска обновляются:

- `data/processed/daily_dashboard_data.json`
- `dist/udp_daily_dashboard.html`

