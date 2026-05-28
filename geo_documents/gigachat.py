from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from dataclasses import dataclass

DEFAULT_GIGACHAT_URL = "http://91.146.28.240:8041/api/gigachat/simple"
DEFAULT_MAX_TOKENS = 100_000
MAX_REPORT_TEXT_CHARS = 150_000

EXPLANATORY_NOTE_TITLE = (
    "1. Пояснительная записка по инженерно-геологическим изысканиям"
)

EXPLANATORY_NOTE_SECTIONS = """
1.1 Введение — объект изысканий, исполнитель, договор/основание работ, уровень ответственности, вид и стадия строительства, цели и задачи изысканий
1.2 Изученность инженерно-геологических условий
1.3 Физико-географические и техногенные условия
1.4 Методика и технология выполнения работ
1.5 Геологическое строение
1.6 Специфические грунты
1.7 Гидрогеологические условия
1.8 Свойства грунтов
1.9 Геологические и инженерно-геологические процессы
1.10 Выводы и рекомендации
  1.10.1 В геоморфологическом отношении
  1.10.2 В геологическом строении
  1.10.3 Гидрогеологические условия (выводы)
  1.10.4 Грунты и их свойства (выводы, ИГЭ)
  1.10.5 Сейсмичность / категория сложности (если есть данные)
  1.10.6 Опасные процессы (если есть данные)
  1.10.7 Категория сложности по СП 47.13330.2016 (если есть данные)
  1.10.8 Рекомендации — нумерованный список: 1) … 2) … 3) …
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
        "Составь пояснительную записку по инженерно-геологическим изысканиям "
        "исключительно на основе приведённых ниже исходных материалов.\n\n"
        "Формат — как в официальном разделе ИГИ «1. Пояснительная записка».\n\n"
        "СТРОГИЕ ПРАВИЛА ДОСТОВЕРНОСТИ (обязательны):\n"
        "1. Используй ТОЛЬКО сведения из блоков «Источник N» ниже. "
        "Запрещено опираться на общие знания, домыслы, типовые шаблоны.\n"
        "2. Каждый абзац с фактами завершай ссылкой на источник: "
        "[ист.: Источник 2] или [ист.: Источник 1; Источник 3].\n"
        "3. Числа, даты, адреса, названия организаций, нормативы (СП, ГОСТ) — "
        "только как в источнике, без округления и «уточнений».\n"
        "4. Если данных для раздела нет — пропусти раздел или одной строкой: "
        "«Сведения в представленных материалах не приведены.»\n"
        "5. Не формулируй выводов и рекомендаций, которых нет в исходных текстах.\n"
        "6. Не используй слова «вероятно», «предположительно», «как правило», "
        "«обычно» — только подтверждённые факты.\n\n"
        "Требования к оформлению:\n"
        f"- Первая строка (заголовок): «{EXPLANATORY_NOTE_TITLE}»\n"
        "- Далее разделы 1.1, 1.2, … 1.10, подразделы 1.10.1–1.10.8\n"
        "- Каждый раздел: строка «1.X Название», затем текст абзацами со ссылками [ист.: …]\n"
        "- В 1.10.8 «Рекомендации» — только рекомендации из источников, список 1), 2), 3)\n"
        "- Без Markdown (#, **). Раздел 1.11 (перечень ГОСТ) не включать\n"
        "- Стиль: официально-деловой, как в инженерном отчёте\n\n"
        "Обязательная структура разделов:\n"
        f"{EXPLANATORY_NOTE_SECTIONS}\n\n"
        f"{sources_block}"
        "Текст исходных материалов отчёта:\n"
        f"{trimmed}"
    )
