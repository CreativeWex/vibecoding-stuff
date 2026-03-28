# Форматирование по ГОСТу

В репозитории лежит **Cursor Skill «ГОСТ Word Formatter»** (`.cursor/skills/gost-word-formatter/`): он задаёт агенту правила и скрипт для приведения документов Word к оформлению по **ГОСТ Р 7.0.11-2008** (поля A4, шрифты Times New Roman, интервалы, заголовки, подписи к рисункам и таблицам, нумерация страниц с второй страницы).

## Запуск форматирования

1. Установите зависимость один раз:

   ```bash
   python3 -m pip install -r .cursor/skills/gost-word-formatter/scripts/requirements.txt
   ```

2. Укажите путь к своему `.docx`:

   ```bash
   python3 .cursor/skills/gost-word-formatter/scripts/gost_format.py "/путь/к/документу.docx"
   ```

   Результат сохранится рядом с исходником как `ГОСТформат<имя>.docx`. Резервная копия **не создаётся**, если явно не указать **`--backup`** (тогда рядом появится `*.gost_backup`). По умолчанию скрипт **добавляет подписи** к рисункам и таблицам без подписи (текст из контекста); отключение: **`--no-infer-captions`**. См. также `--page-number`, `--margin-preset`, `--margins-mm`, `--progress`, `--toc` / `--no-toc` (`python3 .../gost_format.py -h`, `SKILL.md`).

3. Чтобы skill подхватывался в Cursor, папка `gost-word-formatter` должна быть в `.cursor/skills/` проекта или в `~/.cursor/skills/` для всех проектов.

Подробности и триггеры — в `.cursor/skills/gost-word-formatter/SKILL.md`.
