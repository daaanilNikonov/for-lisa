# Центр компетенций 1С-ЭПД — Юлиана Юнусова

Материалы для анонса в чате Центра продаж.

## Файлы

| Файл | Назначение |
|------|------------|
| [`chat_message.md`](./chat_message.md) | Готовый текст сообщения для чата |
| [`Yuliana_Yunusova_1C_EPD_competence_card.pdf`](./Yuliana_Yunusova_1C_EPD_competence_card.pdf) | PDF-карточка в стиле ГК Форус |
| [`Yuliana_Yunusova_1C_EPD_competence_card.png`](./Yuliana_Yunusova_1C_EPD_competence_card.png) | PNG-превью той же карточки |
| `yuliana_photo.jpg` / `yuliana_cutout.png` | Исходное фото и вырезанный портрет |

## Пересборка карточки

```bash
python3 scripts/build_epd_competence_card.py
```

Стиль: тёмный шаблон ГК Форус (`#1A1A1A`, акцент `#26A6E0`), как в `Презентация ГК Форус темный шаблон 16х9`.

Темы обращений взяты из презентации `1С-ЭПД Мастерская для менеджеров` (выбор решения, тарифы/пакеты, сложные кейсы, подпись/МЧД, линия поддержки) плюс возвраты по запросу анонса.
