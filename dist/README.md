# Архив для загрузки на хостинг (SSH)

Файл: `forus-test-kabinet-hosting.zip`

## Загрузка

```bash
scp forus-test-kabinet-hosting.zip user@your-server:/var/www/
ssh user@your-server
cd /var/www
unzip forus-test-kabinet-hosting.zip
# сайт будет в /var/www/test-kabinet/
```

Либо распакуйте содержимое `test-kabinet/` сразу в корень сайта (чтобы `index.html` был в document root).

Внутри архива есть `DEPLOY-VDS.md` с краткой инструкцией.

## Обновление текста на оплаченном домене

См. `UPDATE-SITE.md`. Короткий архив: `forus-code-update.zip` (4 файла без картинок).
