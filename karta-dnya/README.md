# Карта дня · Группа продуктового запуска

Внутренний лендинг ГК Форус для ежедневной работы менеджеров группы продуктового запуска.

**Локальный URL:** [`/karta-dnya`](http://localhost:3847/karta-dnya)

Оформление — тёмная корпоративная тема шаблона «Презентация ГК Форус темный шаблон 16×9».

## Постоянный хостинг (Render)

Туннель Cloudflare временный. Для постоянной ссылки разверните сервис на Render (бесплатно):

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/daaanilNikonov/for-lisa)

1. Войдите на [render.com](https://dashboard.render.com) через GitHub.
2. Нажмите кнопку выше **или** создайте Web Service вручную:
   - Repository: `daaanilNikonov/for-lisa`
   - Root Directory: `karta-dnya`
   - Start Command: `npm start`
   - Plan: Free
3. В Environment добавьте `GITHUB_TOKEN` — fine-grained PAT GitHub с правом **Contents: Read and write** на этот репозиторий.  
   Тогда база `data/db.json` будет сохраняться в GitHub и не пропадёт при «засыпании» бесплатного инстанса.
4. После деплоя ссылка будет вида: `https://forus-karta-dnya.onrender.com/karta-dnya`

На бесплатном тарифе сервис может «засыпать» после простоя ~15 минут — первый заход после сна занимает 30–60 секунд, дальше работает как обычно.

## Возможности

- Вкладка **Все** — обзор чеклистов и досок всех менеджеров
- У каждого менеджера свой чеклист на день и своя доска стикеров
- Утро: создать задачи → Вечер: галочки → в архив только после сохранения вечером (`Имя YYYY-MM-DD`)
- Удаление записей из архива (для тестовых прогонов)
- Переименование менеджера (кнопка или двойной клик по вкладке)
- Данные: [`data/db.json`](./data/db.json) (локально или через GitHub API на хостинге)

## Локальный запуск

```bash
cd karta-dnya
npm start
```

Откройте http://localhost:3847/karta-dnya

Нужен только Node.js 18+.
