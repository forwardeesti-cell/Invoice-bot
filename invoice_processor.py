"""
Обработка PDF счетов через Claude API
"""

import base64
import json
import logging
import os
import re
from pathlib import Path

import anthropic

logger = logging.getLogger(__name__)


class InvoiceProcessor:
    def __init__(self):
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        if not api_key:
            raise ValueError("Не задан ANTHROPIC_API_KEY в .env файле!")
        self.client = anthropic.Anthropic(api_key=api_key)

    async def extract(self, pdf_path: str) -> dict | None:
        """Извлечь данные из PDF через Claude"""
        try:
            # Читать PDF как base64
            with open(pdf_path, 'rb') as f:
                pdf_data = base64.standard_b64encode(f.read()).decode('utf-8')

            prompt = """Ты — эксперт по анализу счетов и накладных. 
Проанализируй этот PDF-документ и извлеки все данные.

Верни ТОЛЬКО JSON (без markdown, без пояснений) в следующем формате:
{
  "number": "номер счёта или null",
  "date": "дата в формате DD.MM.YYYY или null",
  "object": "название объекта/стройки/адреса или null",
  "supplier": "название поставщика или null",
  "items": [
    {
      "name": "название позиции",
      "quantity": число или null,
      "unit": "единица измерения или null",
      "price": цена за единицу как число или null,
      "amount": итоговая сумма позиции как число,
      "category": "категория из списка ниже или 'unknown' если непонятно"
    }
  ],
  "total": общая сумма как число,
  "vat": сумма НДС как число или null,
  "currency": "RUB" или другая валюта
}

Возможные категории для поля category:
- Материалы
- Инструменты и оборудование
- Услуги
- Транспорт и логистика
- Аренда
- Электрика
- Сантехника
- Отделочные работы
- Строительные работы
- Прочее
- unknown (если категория неочевидна)

ВАЖНО: 
- Если объект не указан в документе — поставь null
- Если позиция непонятна — поставь category: "unknown"  
- Все суммы — числа без пробелов и символов валюты
- Верни ТОЛЬКО JSON, никакого другого текста"""

            response = self.client.messages.create(
                model="claude-opus-4-5",
                max_tokens=4000,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "document",
                                "source": {
                                    "type": "base64",
                                    "media_type": "application/pdf",
                                    "data": pdf_data
                                }
                            },
                            {
                                "type": "text",
                                "text": prompt
                            }
                        ]
                    }
                ]
            )

            raw = response.content[0].text.strip()
            
            # Убрать возможные markdown теги
            raw = re.sub(r'^```(?:json)?\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw)
            
            data = json.loads(raw)
            logger.info(f"Извлечено: №{data.get('number')}, позиций: {len(data.get('items', []))}")
            return data

        except json.JSONDecodeError as e:
            logger.error(f"Ошибка парсинга JSON от Claude: {e}\nОтвет: {raw[:500]}")
            return None
        except Exception as e:
            logger.error(f"Ошибка извлечения данных: {e}", exc_info=True)
            return None
