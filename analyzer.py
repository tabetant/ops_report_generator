import json
import os
import re

MOCK_RESPONSE = {
    "kpis": [
        {"name": "Defect Rate", "value": "3.28%", "status": "warning"},
        {"name": "Avg Units Produced", "value": "402.59", "status": "ok"},
        {"name": "Avg Downtime", "value": "0.69 hrs", "status": "ok"},
        {"name": "Max Single-Day Defects", "value": "150", "status": "critical"},
        {"name": "Max Downtime Event", "value": "9.5 hrs", "status": "critical"},
    ],
    "anomalies": [
        {
            "description": "Line L1 on 2024-01-05 reported 150 defects — 12x above average. Likely equipment or material failure.",
            "severity": "critical",
        },
        {
            "description": "Line L2 on 2024-01-06 recorded 9.5 hours of downtime with zero production.",
            "severity": "critical",
        },
        {
            "description": "Three shifts reported zero units produced with no recorded downtime — data may be missing or lines were idle without logging.",
            "severity": "warning",
        },
    ],
    "actions": [
        "Investigate root cause of L1's 150-defect spike on Jan 5 — pull maintenance logs and material batch records.",
        "Review L2's 9.5-hour downtime on Jan 6 — confirm whether the event was logged correctly and assess equipment status.",
        "Audit shifts with zero production and zero downtime — establish a mandatory logging protocol for idle shifts.",
        "Set automated alerts for defect counts exceeding 30 per shift and downtime events exceeding 4 hours.",
    ],
}


def build_prompt(summary: dict) -> str:
    return f"""You are an operations analyst reviewing manufacturing data.

Data summary:
- Rows: {summary["row_count"]}
- Date range: {summary["date_range"]["start"]} to {summary["date_range"]["end"]}
- Lines: {", ".join(summary["lines"])}
- Shifts: {", ".join(summary["shifts"])}
- Defect rate: {summary["defect_rate"]}%
- Units produced — mean: {summary["stats"]["units_produced"]["mean"]}, min: {summary["stats"]["units_produced"]["min"]}, max: {summary["stats"]["units_produced"]["max"]}
- Defects — mean: {summary["stats"]["defects"]["mean"]}, min: {summary["stats"]["defects"]["min"]}, max: {summary["stats"]["defects"]["max"]}
- Downtime hours — mean: {summary["stats"]["downtime_hours"]["mean"]}, min: {summary["stats"]["downtime_hours"]["min"]}, max: {summary["stats"]["downtime_hours"]["max"]}

Return ONLY valid JSON (no markdown, no explanation) matching this exact structure:
{{
  "kpis": [
    {{"name": "string", "value": "string", "status": "ok|warning|critical"}}
  ],
  "anomalies": [
    {{"description": "string", "severity": "ok|warning|critical"}}
  ],
  "actions": ["string"]
}}

Identify 4-6 KPIs, flag 2-4 anomalies, and provide 3-5 concrete recommended actions."""


def parse_response(raw: str) -> dict:
    # Strip markdown code fences if present
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw.strip(), flags=re.MULTILINE)
    cleaned = re.sub(r"\s*```$", "", cleaned.strip(), flags=re.MULTILINE)
    return json.loads(cleaned.strip())


def call_claude(prompt: str, mock: bool = True) -> dict:
    if mock:
        return MOCK_RESPONSE

    api_key = (os.getenv("ANTHROPIC_API_KEY") or "").strip()
    if not api_key:
        raise EnvironmentError(
            "ANTHROPIC_API_KEY is not set. Add it to your .env file."
        )

    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text
    return parse_response(raw)


if __name__ == "__main__":
    import json as _json
    from pathlib import Path
    import loader

    sample_path = Path(__file__).parent / "samples" / "operations_sample.csv"
    df = loader.load_file(str(sample_path))
    loader.validate_columns(df, loader.REQUIRED_COLUMNS)
    summary = loader.summarize(df)

    prompt = build_prompt(summary)
    print("--- Prompt (first 300 chars) ---")
    print(prompt[:300], "...\n")

    result = call_claude(prompt, mock=True)
    print("--- Mock Analysis ---")
    print(_json.dumps(result, indent=2))
