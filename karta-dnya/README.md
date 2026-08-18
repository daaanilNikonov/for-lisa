# Карта дня · Группа продуктового запуска

Внутренний лендинг ГК Форус для ежедневной работы менеджеров группы продуктового запуска.

**Локальный URL:** [`/karta-dnya`](http://localhost:3847/karta-dnya)

Оформление — тёмная корпоративная тема шаблона «Презентация ГК Форус темный шаблон 16×9».

## Постоянный хостинг (Render)

Важно: лендинг пока лежит в ветке  
`cursor/karta-dnya-produktovogo-zapuska-ed4c`  
(в `main` его ещё нет). Поэтому кнопка «Deploy to Render» с `main` может показать **Not Found**.

### Как задеплоить вручную (работает)

1. Войдите на [dashboard.render.com](https://dashboard.render.com) через GitHub.
2. **New → Web Service** → выберите репозиторий `daaanilNikonov/for-lisa`.
3. Укажите:
   - **Branch:** `cursor/karta-dnya-produktovogo-zapuska-ed4c`
   - **Root Directory:** `karta-dnya`
   - **Runtime:** Node
   - **Build Command:** можно оставить пустым или `node -v`
   - **Start Command:** `npm start`
   - **Instance type:** Free
4. В **Environment** добавьте:
   - `GITHUB_TOKEN` — fine-grained PAT с **Contents: Read and write**
   - `GITHUB_REPO` = `daaanilNikonov/for-lisa`
   - `GITHUB_BRANCH` = `cursor/karta-dnya-produktovogo-zapuska-ed4c`
   - `GITHUB_DB_PATH` = `karta-dnya/data/db.json`
5. Deploy. Ссылка будет вида:  
   `https://<имя-сервиса>.onrender.com/karta-dnya`

На бесплатном тарифе сервис может «засыпать» после простоя ~15 минут — первый заход после сна занимает 30–60 секунд.

Чтобы кнопка Deploy to Render заработала без выбора ветки — смержите эту ветку в `main`.

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
