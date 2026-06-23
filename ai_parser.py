import anthropic
import json
import re

SYSTEM_PROMPT = """You are an expert at analyzing marketing documents (audits, strategies) and structuring them into presentation slides.

Your output must be ONLY valid JSON — no markdown, no explanation, no code blocks.

JSON structure:
{
  "cover": {
    "client_name": "string",
    "platform": "string (e.g. Google Ads, Meta Ads)",
    "document_type": "Audit or Strategy",
    "document_title": "string",
    "subtitle": "string",
    "year": "string (4 digits)"
  },
  "doc_label": "string (e.g. 'Audit | Google Ads')",
  "sections": [
    {
      "number": "01",
      "title": "string",
      "bullets": [
        {"label": "Key term:", "text": "Explanation text"}
      ]
    }
  ]
}

Rules:
- 5 to 13 sections
- 2 to 5 bullets per section
- Keep all text in the SAME LANGUAGE as the source document
- Labels end with a colon
- Text is concise but complete
- Number sections starting at "01", zero-padded"""


def parse_document(text: str) -> dict:
    client = anthropic.Anthropic()

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"Parse this document into presentation slides:\n\n{text}"
            }
        ]
    )

    raw = message.content[0].text.strip()

    # Strip markdown code blocks if model wraps them
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)

    return json.loads(raw)
