ГОТОВЫЙ КОМПЛЕКТ ДЛЯ RENDER + ЕЖЕДНЕВНОЙ ОТПРАВКИ 07:00

Файлы:
- bot_render.py — Telegram-бот через webhook.
- requirements.txt — зависимости.
- render.yaml — настройки Render.
- daily_sender.py — отправка расписания один раз (используется GitHub Actions).
- .gitignore

ВАЖНО: токен не хранится в коде. В Render/GitHub его нужно добавить как BOT_TOKEN.

АВТО 07:00:
GitHub Actions должен запускать daily_sender.py ежедневно в 00:00 UTC, что соответствует 07:00 Asia/Barnaul.
Если нужен другой часовой пояс, поменяй cron в workflow.

AUTO_CHAT_ID:
1. После деплоя напиши боту /myid.
2. Бот покажет chat_id.
3. В Render добавь AUTO_CHAT_ID = это число.
4. В GitHub Secrets добавь BOT_TOKEN и AUTO_CHAT_ID.
