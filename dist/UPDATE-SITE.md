# Обновить какойтысотрудник.рф — текст и Яндекс.Форма

На сервере сейчас **старая** версия. В репозитории уже есть нужные правки:

- в 1 вопросе: **главный бухгалтер** (вместо «глб»)
- список КЭДО только:
  - 1С:Кабинет сотрудника
  - Nopaper
  - Топ Фактор
  - Directum
  - Другое
- выбранный КЭДО автоматически подставляется в поле «КЭДО» Яндекс.Формы
- вопрос про справки: «Как часто сотрудники приходят к вам за справками»

## Быстрый способ (только код, ~24 КБ)

Скачайте: [forus-code-update.zip](https://github.com/daaanilNikonov/for-lisa/raw/cursor/quiz-landing-scripts-55dd/dist/forus-code-update.zip)

Распакуйте **в ту же папку**, где лежит текущий `index.html` на сервере (заменить 4 файла):

- `index.html`
- `styles.css`
- `app.js`
- `scripts-data.js`

Через файловый менеджер хостинга / SFTP / SCP — перезаписать эти файлы.

## Полный пакет (код + картинки)

Если ещё нет папки `assets/` (картинки не открываются):

[forus-test-kabinet-hosting.zip](https://github.com/daaanilNikonov/for-lisa/raw/cursor/quiz-landing-scripts-55dd/dist/forus-test-kabinet-hosting.zip)

Распаковать содержимое `test-kabinet/` в корень сайта.

## Проверка после выкладки

1. Жёсткое обновление страницы: Ctrl+Shift+R  
2. Вопрос 1 — пункт «главный бухгалтер»  
3. Ответ «есть КЭДО» → список из 5 пунктов (с «1С:Кабинет сотрудника»)  
4. В заявке выбранный КЭДО уходит в Яндекс.Форму  
5. `http://какойтысотрудник.рф/assets/results/detoks.jpg` открывается как картинка

## Патч: «расчетными листками» (вопрос 3)

Скачайте и замените `scripts-data.js` в корне сайта:

- [forus-patch-raschetnymi.zip](https://github.com/daaanilNikonov/for-lisa/raw/cursor/quiz-landing-scripts-55dd/dist/forus-patch-raschetnymi.zip) — только `scripts-data.js`
- или полный [forus-code-update.zip](https://github.com/daaanilNikonov/for-lisa/raw/cursor/quiz-landing-scripts-55dd/dist/forus-code-update.zip)

Было: «Как вы работаете с расчетными и приказами»  
Стало: «Как вы работаете с расчетными листками и приказами»
