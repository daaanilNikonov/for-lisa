# Почему нет картинок на какойтысотрудник.рф

Сервер на запросы вроде `/assets/results/detoks.jpg` отдаёт `index.html`
вместо файла картинки. Обычно это значит:

1. папка `assets/` не загружена на хостинг, или
2. загружена не в ту директорию (не рядом с `index.html`).

## Быстрый фикс

Рядом с `index.html` на сервере должна быть структура:

```
index.html
styles.css
app.js
scripts-data.js
assets/
  brand/logo-forus.png
  results/*.jpg
```

### Вариант A — только картинки

```bash
# локально
scp forus-assets-only.zip user@SERVER:/path/to/site/

# на сервере
cd /path/to/site
unzip -o forus-assets-only.zip
# появится ./assets/...
```

### Вариант B — перезалить весь лендинг

```bash
scp forus-test-kabinet-hosting.zip user@SERVER:/tmp/
ssh user@SERVER
cd /path/to/site
unzip -o /tmp/forus-test-kabinet-hosting.zip
# если распаковалось в test-kabinet/ — перенесите содержимое на уровень index.html:
cp -a test-kabinet/. ./
```

Проверка после загрузки:

```bash
curl -I https://какойтысотрудник.рф/assets/results/detoks.jpg
# Content-Type должен быть image/jpeg, НЕ text/html
```
