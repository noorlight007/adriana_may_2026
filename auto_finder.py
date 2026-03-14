from typing import Any, Dict
import re

def return_Schema() -> Dict[str, Any]:
    SCHEMA: Dict[str, Any] = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "email": {
                "type": "string",
                "description": 'Company or contact email. Return "null" if not found.'
            }
        },
        "required": ["email"]
    }
    return SCHEMA

def get_all_instructions():
    SYSTEM_INSTRUCTIONS = (
        "You are an information extraction engine.\n"
        "Your task is to explore the given website and extract only one email address.\n"
        "- Extract the most relevant company email or contact person's work email.\n"
        "- Only extract an email if it is explicitly present in the provided content.\n"
        "- Do NOT invent, guess, or infer any email address.\n"
        "- If no valid email is found, return the literal string \"null\".\n"
        "- Return ONLY raw JSON matching the required schema.\n"
        "- Do not return markdown fences, explanations, or extra text.\n"
        "- Output format must be exactly like:\n"
        "{\n"
        "  \"email\": \"example@company.com\"\n"
        "}\n"
        "- If no email exists, output:\n"
        "{\n"
        "  \"email\": \"null\"\n"
        "}\n"
    )

    USER_INSTRUCTIONS_TEMPLATE = (
        "Extract the single most relevant email address by exploring the following website link.\n"
        "If no email is present, return \"null\".\n"
        "Return ONLY raw JSON matching the required schema.\n\n"
        "CONTENT START\n"
        "{content}\n"
        "CONTENT END\n"
    )

    return SYSTEM_INSTRUCTIONS, USER_INSTRUCTIONS_TEMPLATE


def strip_code_fences(s: str) -> str:
    """Remove markdown code fences like ``` or ```json."""
    s = s.strip()
    if s.startswith("```"):
        lines = s.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        s = "\n".join(lines).strip()
    return s


def find_json_block(s: str) -> str:
    """If JSON isn't clean, try to extract the largest {...} or [...] block."""
    s = s.strip()
    obj_match = re.search(r"\{.*\}\s*$", s, flags=re.DOTALL)
    if obj_match:
        return obj_match.group(0)
    arr_match = re.search(r"\[.*\]\s*$", s, flags=re.DOTALL)
    if arr_match:
        return arr_match.group(0)
    return s


def normalize_nulls(data: Dict[str, Any]) -> Dict[str, Any]:
    """Force missing/empty email to the literal string 'null'."""
    def to_str_null(v: Any) -> str:
        if v is None:
            return "null"
        if isinstance(v, str):
            return v.strip() if v.strip() else "null"
        return str(v)

    data["email"] = to_str_null(data.get("email"))
    return data