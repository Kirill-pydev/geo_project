from __future__ import annotations

import json
import urllib.error
import urllib.request
from dataclasses import dataclass

DEFAULT_GIGACHAT_URL = "http://91.146.28.240:8041/api/gigachat/simple"
DEFAULT_MAX_TOKENS = 100_000
MAX_REPORT_TEXT_CHARS = 150_000

EXPLANATORY_NOTE_TITLE = (
    "2. Пояснительная записка по инженерно-геологическим изысканиям"
)

# Внутренний чек-лист для модели (не выводить в текст записки дословно).
_EXPLANATORY_NOTE_CHECKLIST = """
1.1 Введение: наименование объекта; исполнитель; цели вида работ; цели и задачи
1.2 Изученность: материалы ранее выполненных изысканий
1.3 Физико-географические и техногенные условия: административное и геоморфологическое
  положение; климат; снеговой район по СП 20; нормативная глубина промерзания; техногенные условия
1.4 Методика: категория сложности; число скважин; вид привязки; способ бурения; отбор проб;
  испытания образцов; объёмы работ
1.5 Геологическое строение: глубина бурения; состав разреза; основание выделения ИГЭ; изменение мощностей ИГЭ
1.6 Специфические грунты: виды; скважины вскрытия; мощности
1.7 Гидрогеологические условия: горизонты; тип и породы; химсостав вод; подтопляемость по СП 11; прогноз УГВ
1.8 Свойства грунтов: перечень ИГЭ; границы; физико-механические свойства; пучинистость; коррозионная агрессивность
1.9 Процессы: опасные процессы; сейсмичность; суффозия
1.10 Контроль и приёмка: внутренний и наружный контроль
1.11 Заключение: краткие выводы по разделам 1.1–1.10; самостоятельные рекомендации по строительству
""".strip()

OUTPUT_FORMAT_EXAMPLE = """
2. Пояснительная записка по инженерно-геологическим изысканиям

@Р@ 1.1 Введение

@П@ Наименование объекта
Текст абзаца. [ист.: Источник 1]

@П@ Исполнитель
Текст абзаца. [ист.: Источник 1]

@Р@ 1.11 Заключение

@П@ Краткие выводы
Текст по разделам 1.1–1.10. [ист.: Источник 2]

@П@ Рекомендации
1) Рекомендация, сформулированная на основе фактов отчёта.
2) Следующая рекомендация.
""".strip()


@dataclass
class GigaChatResponse:
    content: str
    used_tokens: int


def call_simple(
    text: str,
    *,
    url: str = DEFAULT_GIGACHAT_URL,
    max_tokens: int = DEFAULT_MAX_TOKENS,
    timeout_sec: int = 300,
) -> GigaChatResponse:
    """Отправляет текст на /api/gigachat/simple и возвращает ответ модели."""
    payload = {"text": text, "maxTokens": max_tokens}
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=timeout_sec) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"GigaChat HTTP {exc.code}: {body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"GigaChat недоступен: {exc.reason}") from exc

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Некорректный JSON от GigaChat: {raw[:500]}") from exc

    content = str(parsed.get("content", "")).strip()
    used_tokens = int(parsed.get("usedTokens") or 0)

    if not content:
        raise RuntimeError("GigaChat вернул пустой ответ.")
    if content.startswith("Ошибка при обращении"):
        raise RuntimeError(content)

    return GigaChatResponse(content=content, used_tokens=used_tokens)


def build_explanatory_note_prompt(
    report_text: str,
    source_labels: list[str] | None = None,
) -> str:
    """Формирует промпт для пояснительной записки в формате типового ИГИ."""
    trimmed = report_text.strip()
    if len(trimmed) > MAX_REPORT_TEXT_CHARS:
        trimmed = trimmed[:MAX_REPORT_TEXT_CHARS] + "\n\n[... текст сокращён ...]"

    sources_block = ""
    if source_labels:
        sources_block = (
            "Допустимые источники (только они, других нет):\n"
            + "\n".join(f"- {label}" for label in source_labels)
            + "\n\n"
        )

    return (
        "Составь пояснительную записку по инженерно-геологическим изысканиям.\n"
        "Записка вставляется в отчёт сразу после «Содержания» как раздел 2.\n\n"
        "ПРАВИЛА ДОСТОВЕРНОСТИ (разделы 1.1–1.10):\n"
        "1. Факты, цифры, адреса, нормативы — ТОЛЬКО из блоков «Источник N».\n"
        "2. Каждый абзац с фактами завершай: [ист.: Источник 2] или [ист.: Источник 1; Источник 3].\n"
        "3. Если данных нет — одна строка: «Сведения в представленных материалах не приведены.»\n"
        "4. Не используй «вероятно», «предположительно», «обычно».\n"
        "5. Пиши подробно, но без воды.\n\n"
        "РАЗДЕЛ 1.11 — ОСОБЫЕ ПРАВИЛА:\n"
        "• «Краткие выводы» — только факты из источников, по разделам 1.1–1.10.\n"
        "• «Рекомендации» — СФОРМУЛИРУЙ САМ (не копируй из исходников дословно): "
        "3–7 конкретных пунктов «1) … 2) …» на основе выводов и фактов отчёта. "
        "Не задавай вопросов — только готовые рекомендации по строительству/проектированию.\n\n"
        "ФОРМАТ ВЫВОДА (строго соблюдай, это готовый текст отчёта, не черновик):\n"
        f"- Первая строка: «{EXPLANATORY_NOTE_TITLE}»\n"
        "- Раздел (пункт): «@Р@ 1.X Краткое название» — одна строка, без «?» и без нумерованных вопросов\n"
        "- Подпункт: «@П@ Краткий подзаголовок» — 2–5 слов (например «Наименование объекта», "
        "«Исполнитель», «Климатические условия»), не «Кто является исполнителем?»\n"
        "- После подзаголовка — текст абзаца(ов)\n"
        "- Между разделами — пустая строка\n"
        "- Не используй формулировки чек-листа («ответь», «укажи», «имеются ли»)\n"
        "- Без Markdown (#, **)\n"
        "- Не включай чек-лист и этот промпт в ответ\n\n"
        "Пример оформления:\n"
        f"{OUTPUT_FORMAT_EXAMPLE}\n\n"
        "Содержание разделов (раскрой все темы):\n"
        f"{_EXPLANATORY_NOTE_CHECKLIST}\n\n"
        f"{sources_block}"
        "Текст исходных материалов отчёта:\n"
        f"{trimmed}"
    )
