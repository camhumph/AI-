import argparse
import json
import subprocess
from pathlib import Path


BASE = Path(r"C:\CMS_AI\geometry_classifier")
DATA = BASE / "data" / "geometry_ai_examples.jsonl"
KNOWLEDGE = BASE / "mold_geometry_knowledge.md"


def load_examples():
    if not DATA.exists():
        raise SystemExit(f"Missing {DATA}. Run collect_geometry_dataset.py first.")
    examples = []
    with DATA.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                examples.append(json.loads(line))
    return examples


def find_example(examples, job_text):
    job_text = job_text.lower()
    for ex in examples:
        if job_text in ex.get("job", "").lower() or job_text in ex.get("source", "").lower():
            return ex
    return None


def build_prompt(example):
    knowledge = KNOWLEDGE.read_text(encoding="utf-8")
    rows = example["rows"]
    compact = json.dumps(rows, indent=2)
    return f"""
You are a mold-base geometry interpreter for CMS quoting.

Rules/knowledge:
{knowledge}

Task:
Classify the CAD rows below. Exact shop-standard name tokens (A-PLATE, B-PLATE,
SC-RETAINER, SC-BACKUP, EJ-RET, EJ-BACKUP, RAIL, LDR-PIN, LBB, PLC75/LATCH-LOCK/
SAFETY-STRAP) are strong anchor evidence; only fall back to geometry when names
are generic or missing.
Return a concise table with:
Index, Role, Confidence, Geometry Reason.

Also list:
- full-footprint stack order
- rails
- pin plate / ejector plate
- leader pins
- bushings
- support pillars
- likely parting line

Job: {example["job"]}
Source: {example["source"]}
Summary: {json.dumps(example["summary"])}

Rows:
{compact}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("job", help="Job text to search, for example C18635 or 8418")
    ap.add_argument("--model", default="qwen3.5:4b", help="Ollama model name")
    args = ap.parse_args()

    examples = load_examples()
    ex = find_example(examples, args.job)
    if not ex:
        print(f"No likely standard/non-BMS example matched: {args.job}")
        print("Available examples:")
        for item in examples[:50]:
            print("  " + item["job"])
        raise SystemExit(1)

    prompt = build_prompt(ex)
    cmd = ["ollama", "run", args.model]
    try:
        result = subprocess.run(
            cmd,
            input=prompt,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=300,
        )
    except FileNotFoundError:
        raise SystemExit("Ollama is not installed or not on PATH.")

    if result.stderr.strip():
        print(result.stderr.strip())
    print(result.stdout)


if __name__ == "__main__":
    main()
