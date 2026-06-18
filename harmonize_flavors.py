#!/usr/bin/env python3
"""
harmonize_flavors.py — Stage 2c (LLM-assisted harmonization, human-validated).

Goal (PRD §5.3 / TECHNICAL §11): close the harmonization gaps where `flavor_en`
falls back to a title-cased local label because the local term was not in the
`flavor_map` of Stage 1. The same pattern is reusable for weight/format.

Pipeline:
  1. Read the canonical parquet, list the DISTINCT unmapped flavor labels per
     country (heuristic: labels that are not already canonical English).
  2. Ask the Anthropic API to propose `local -> EN canonical` mappings, grouped
     by country, returning STRICT JSON.
  3. YOU validate (review the printed proposal / the written JSON file).
  4. Paste the approved entries into the per-country `flavor_map` of
     build_category_db.py and re-run Stage 1.

This script never writes the parquet and never edits Stage 1 automatically:
the human stays in the loop, by design.

Usage:
  export ANTHROPIC_API_KEY=sk-ant-...
  python3 harmonize_flavors.py category_db.parquet [--country FR] [--apply-dry]

Requires: pandas, pyarrow, anthropic
"""
import argparse, json, os, sys
import pandas as pd

MODEL = "claude-opus-4-8"  # latest, most capable for this mapping task

# Canonical EN set already produced by Stage 1 (extend as the map grows).
CANONICAL_EN = {
    "Plain", "Vanilla", "Lemon", "Coconut", "Strawberry", "Blueberry",
    "Raspberry", "Blackberry", "Cherry", "Pineapple", "Mango", "Orange",
    "Passion Fruit", "Peach", "Mixed", "Various", "Stracciatella",
    "Chocolate", "Hazelnut", "Coffee", "Mixed Berries", "Cream",
    "Apple Cinnamon", "Honey Walnut", "Sachertorte", "Choco Coconut",
    "Walnut", "Citrus", "Fruit Mix", "Sweetened Plain", "Unsweetened Plain",
    "Mandarin",
}

PROMPT = """You harmonize yoghurt flavor labels to a canonical English vocabulary.

For each local label below (country: {country}), return the best canonical English
flavor. Prefer reusing one of these canonical terms when it fits:
{canon}

Rules:
- Compound/rare flavors: give a concise English title-case label (e.g.
  "Frutti di Bosco" -> "Mixed Berries", "Mela Cannella" -> "Apple Cinnamon").
- If a label is clearly NOT a flavor (a brand, a size, noise), map it to null.
- Output STRICT JSON only: an object {{"local label": "EN canonical or null", ...}}.

Local labels:
{labels}
"""


def unmapped_flavors(df, country=None):
    """Distinct flavor_en values that look un-harmonized (not in CANONICAL_EN)."""
    sub = df[df["level"] == "flavor"]
    if country:
        sub = sub[sub["country"] == country]
    out = {}
    for c, g in sub.groupby("country"):
        labels = sorted({str(x) for x in g["flavor_en"].dropna().unique()
                         if str(x) not in CANONICAL_EN})
        if labels:
            out[c] = labels
    return out


def propose(country, labels):
    """Call the Anthropic API for one country's labels. Returns dict local->EN|None."""
    try:
        import anthropic
    except ImportError:
        sys.exit("pip install anthropic  (and set ANTHROPIC_API_KEY)")
    client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY
    msg = client.messages.create(
        model=MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": PROMPT.format(
            country=country,
            canon=", ".join(sorted(CANONICAL_EN)),
            labels="\n".join(f"- {l}" for l in labels),
        )}],
    )
    text = "".join(b.text for b in msg.content if getattr(b, "type", None) == "text")
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON found in model reply:\n{text}")
    return json.loads(text[start:end + 1])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("parquet", nargs="?", default="category_db.parquet")
    ap.add_argument("--country", help="limit to one country code (FR/UK/DE/ES/IT)")
    ap.add_argument("--apply-dry", action="store_true",
                    help="only list unmapped labels; do NOT call the API")
    ap.add_argument("--out", default="flavor_proposals.json")
    args = ap.parse_args()

    df = pd.read_parquet(args.parquet)
    gaps = unmapped_flavors(df, args.country)
    if not gaps:
        print("No unmapped flavors found — nothing to harmonize.")
        return

    for c, labels in gaps.items():
        print(f"\n[{c}] {len(labels)} unmapped labels:")
        for l in labels:
            print(f"   {l}")

    if args.apply_dry:
        print("\n--apply-dry: skipping API calls.")
        return
    if not os.getenv("ANTHROPIC_API_KEY"):
        sys.exit("\nSet ANTHROPIC_API_KEY to request proposals (or use --apply-dry).")

    proposals = {}
    for c, labels in gaps.items():
        print(f"\nRequesting proposals for {c} …")
        proposals[c] = propose(c, labels)
        for local, en in proposals[c].items():
            print(f"   {local!r:30} -> {en!r}")

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(proposals, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {args.out}. REVIEW IT, then fold approved entries into the "
          f"per-country flavor_map of build_category_db.py and re-run Stage 1.")


if __name__ == "__main__":
    main()
