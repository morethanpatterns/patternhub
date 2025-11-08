#!/usr/bin/env python3
"""
Generate PatternHub/manifest.json automatically.

File naming rule:
Author/Block/<author>_<block>_<description>_v<version>.jsx
Example:
  Aldrich/Bodice/aldrich_bodice_close_fitting_v1.jsx
  Hofenbitzer/Trouser/hofenbitzer_trouser_ergonomic_v2.jsx
  Muller/Sleeve/muller_sleeve_two_piece_v3.jsx

Optional sidecar files beside each .jsx:
  - same name + .json (for notes/overrides)
  - same name + .txt  (for plain-text notes)
"""

import json, re
from pathlib import Path
from datetime import date

ROOT = Path(__file__).resolve().parents[1]       # repo root
MANIFEST_PATH = ROOT / "manifest.json"

IGNORE_TOP = {"tools", ".git", ".github", "__pycache__", "docs"}

# Match "_v1.jsx", "_v2.jsx", "_v3.jsx"
FILENAME_RE = re.compile(r"^(?P<slug>.+?)_v(?P<version>\d+)\.jsx$", re.IGNORECASE)


# ──────────────────────────────────────────────────────────────────────────────
def title_case_slug(slug: str) -> str:
    """Turn 'aldrich_bodice_close_fitting' → 'Bodice Close Fitting'."""
    parts = slug.split("_")
    if len(parts) > 1:
        parts = parts[1:]  # drop author's surname
    return " ".join(p.capitalize() for p in parts)


def read_sidecar(jsx_path: Path) -> dict:
    """Read optional .json or .txt next to the JSX file."""
    base = jsx_path.with_suffix("")
    data: dict = {}
    json_path = base.with_suffix(".json")
    txt_path = base.with_suffix(".txt")

    if json_path.exists():
        try:
            with json_path.open("r", encoding="utf-8") as f:
                j = json.load(f)
            for k in ("label", "notes", "requires", "icon"):
                if k in j:
                    data[k] = j[k]
        except Exception as e:
            print(f"⚠️  Could not parse {json_path}: {e}")

    elif txt_path.exists():
        try:
            notes = txt_path.read_text(encoding="utf-8").strip()
            if notes:
                data["notes"] = notes
        except Exception as e:
            print(f"⚠️  Could not read {txt_path}: {e}")

    return data


def collect_versions(block_dir: Path) -> list[dict]:
    entries = []
    for jsx in sorted(block_dir.glob("*.jsx"), key=lambda x: x.name.lower()):
        m = FILENAME_RE.match(jsx.name)
        if not m:
            continue
        slug = m.group("slug")
        version = m.group("version")

        entry = {
            "version": version,
            "status": "stable",
            "path": jsx.relative_to(ROOT).as_posix(),
            "label": title_case_slug(slug),
            "requires": {"illustrator": ">=25"}
        }

        icon_guess = jsx.with_suffix(".png")
        if icon_guess.exists():
            entry["icon"] = icon_guess.relative_to(ROOT).as_posix()

        # merge sidecar info
        overrides = read_sidecar(jsx)
        if "label" in overrides:
            entry["label"] = overrides["label"].strip()
        if "notes" in overrides:
            entry["notes"] = overrides["notes"]
        if "requires" in overrides and isinstance(overrides["requires"], dict):
            req = dict(entry["requires"])
            req.update(overrides["requires"])
            entry["requires"] = req
        if "icon" in overrides:
            entry["icon"] = overrides["icon"]

        entries.append(entry)

    # numeric sort newest → oldest
    entries.sort(key=lambda v: int(v["version"]), reverse=True)
    return entries


def collect_blocks(author_dir: Path) -> list[dict]:
    blocks = []
    for block_dir in sorted([p for p in author_dir.iterdir() if p.is_dir()],
                            key=lambda x: x.name.lower()):
        versions = collect_versions(block_dir)
        if versions:
            blocks.append({
                "id": block_dir.name.lower(),
                "title": f"{author_dir.name} {block_dir.name} Block",
                "category": block_dir.name,
                "versions": versions
            })
    blocks.sort(key=lambda b: b["id"])
    return blocks


def collect_authors() -> list[dict]:
    authors = []
    for author_dir in sorted([p for p in ROOT.iterdir()
                              if p.is_dir() and author_dir_ok(p)], key=lambda x: x.name.lower()):
        blocks = collect_blocks(author_dir)
        if blocks:
            authors.append({
                "id": author_dir.name.lower(),
                "name": author_dir.name.replace("_", " "),
                "blocks": blocks
            })
    return authors


def author_dir_ok(p: Path) -> bool:
    return p.name not in IGNORE_TOP and not p.name.startswith(".")


# ──────────────────────────────────────────────────────────────────────────────
def main():
    manifest = {
        "name": "PatternHub",
        "version": "1.0.0",
        "schema": "https://patternhub.dev/schema/v1",
        "updated": date.today().isoformat(),
        "authors": collect_authors()
    }
    with MANIFEST_PATH.open("w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print(f"✅ Manifest updated → {MANIFEST_PATH}")


if __name__ == "__main__":
    main()
