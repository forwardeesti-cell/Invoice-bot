# 🤖 Invoice Bot — Бот для учёта счетов

Telegram-бот который читает PDF счета, обновляет Excel файлы объектов,
следит за перерасходом материалов и изменением цен.

---

## 📁 Структура файлов

```
invoice_bot/
├── bot.py                 # Основной файл бота
├── invoice_processor.py   # Извлечение данных из PDF (Claude AI)
├── excel_updater.py       # Обновление Excel файлов объектов
├── excel_exporter.py      # Аналитика в Excel
├── database.py            # База данных SQLite
├── requirements.txt       # Зависимости Python
├── .env.example           # Пример настроек
└── objects/               # Папка с Excel файлами объектов (создаётся автоматически)
```

---

## ⚙️ Установка

### 1. Установить Python 3.10+
https://www.python.org/downloads/

### 2. Создать папку и скопировать файлы
```bash
mkdir invoice_bot
cd invoice_bot
# скопируйте все файлы сюда
```

### 3. Установить зависимости
```bash
pip install -r requirements.txt
```

### 4. Создать бота в Telegram
1. Напишите @BotFather в Telegram
2. Команда /newbot
3. Введите имя и username бота
4. Скопируйте токен

### 5. Получить ключ Anthropic API
1. Зайдите на https://console.anthropic.com
2. API Keys → Create Key
3. Скопируйте ключ

### 6. Настроить .env файл
```bash
cp .env.example .env
```
Откройте `.env` и вставьте ваши токены:
```
TELEGRAM_BOT_TOKEN=1234567890:ABCdef...
ANTHROPIC_API_KEY=sk-ant-...
```

### 7. Запустить бота
```bash
# Windows:
python bot.py

# Linux/Mac:
python3 bot.py
```

---

## 🚀 Использование

### Первый запуск:
1. Напишите боту `/start`
2. Используйте `/setxlsx` — отправьте ваш Excel файл объекта
3. Укажите название объекта (например `Tartu mnt`)

### Загрузка счёта:
1. Отправьте PDF файл счёта прямо в чат
2. Если объект не указан в счёте — бот спросит
3. Если раздел непонятен — бот предложит выбрать
4. Бот обновит Excel и пришлёт его обратно

### Команды:
- `/objects` — список всех объектов
- `/report` — краткая аналитика по объекту
- `/export` — полная аналитика в Excel (все счета, история цен)
- `/setxlsx` — привязать/обновить Excel файл объекта

---

## 📊 Что обновляется в Excel автоматически

**Лист Kokku (итоги):**
- Добавляется новая строка в нужный месяц
- Поставщик, номер счёта, дата, сумма материалов

**Листы спецификаций (Kanalisatsioon, Vesi, MÄRG TORU, Sadevee):**
- Обновляется колонка "Закупка" (+количество из счёта)
- Обновляется "Остаток по спецификации"
- Обновляется дата внесения данных
- 🔴 Красный цвет — если остаток < 0 (перерасход)
- 🟠 Оранжевый — если остаток < 10% от нормы
- 🟢 Зелёный — всё в порядке

---

## 🔔 Уведомления бота

| Уведомление | Условие |
|-------------|---------|
| 🔴 Перерасход по позиции | Закуплено больше чем в спецификации |
| ⚠️ Бюджет > 80% | Потрачено более 80% от суммы договора |
| 🚨 Бюджет превышен | Потрачено более 100% |
| 📈 Цена выросла | Цена позиции выше чем в прошлом счёте |
| 📉 Цена упала | Цена позиции ниже чем в прошлом счёте |

---

## 🖥️ Запуск как фоновый сервис (Linux/VPS)

```bash
# Создать systemd сервис
sudo nano /etc/systemd/system/invoicebot.service
```

```ini
[Unit]
Description=Invoice Telegram Bot
After=network.target

[Service]
WorkingDirectory=/home/user/invoice_bot
EnvironmentFile=/home/user/invoice_bot/.env
ExecStart=/usr/bin/python3 bot.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable invoicebot
sudo systemctl start invoicebot
sudo systemctl status invoicebot
```

---

## ❓ Частые вопросы

**Бот не читает PDF**
— Убедитесь что файл является настоящим PDF, а не картинкой с расширением .pdf

**Excel не обновляется**
— Проверьте что файл привязан через /setxlsx
— Проверьте что файл не открыт в другой программе

**Ошибка ANTHROPIC_API_KEY**
— Проверьте .env файл, ключ должен начинаться с sk-ant-
