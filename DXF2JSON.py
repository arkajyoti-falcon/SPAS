#!/usr/bin/env python3
"""
DXF -> Component count tables (top-level + exploded) + feedline/residual analysis.

What it does:
1) TABLE 1: Counts modelspace INSERT blocks (top-level components placed in the DXF).
2) TABLE 2: Explodes nested INSERTs inside BLOCK definitions recursively and counts them too
   (INCLUDE_CONTAINERS mode since LEAF_ONLY=False).
3) Prints breakdown of key blocks (feedline + infeed).
4) Prints FEEDLINE composition (per 1), FEEDLINE total contribution (x6), and RESIDUAL
   (exploded minus feedline contribution). This helps isolate bag-induct / other systems.

Setup:
  pip install ezdxf

Run:
  python DXF2JSON.py
"""

import re
from collections import Counter
from pathlib import Path
import ezdxf

# =============================================================================
# CONFIG
# =============================================================================
INPUT_DXF = r"2211171070_0_Noon_Riyadh_Loop CBS_20K(Rev08).dxf"

# Print only top N (0 = all)
TOP_N_TABLE1 = 0
TOP_N_TABLE2 = 0

# Exclude anonymous/noise blocks
INCLUDE_NOISE = False

# Exclude blocks whose names start with these prefixes (case-insensitive)
EXCLUDE_PREFIXES = [
    "FAL_BLK_ATT_",   # attribute/annotation blocks
    "ACAD_",          # autocad internal
    "A$C",            # autocad internal
]

# Exclude exact block names (case-insensitive)
EXCLUDE_EXACT = [
    "Operator",
    "_InsertPositionStart_",
    "_InsertPositionEnd_",
]

# Sort mode: "count" or "name"
SORT_BY = "count"

# If True, exploded table counts only LEAF blocks (blocks that don’t contain other non-noise inserts)
# For your validation you want containers too, so keep False.
LEAF_ONLY = False

# Which parent assemblies to print child breakdown for
TARGET_BLOCKS = [
    "FAL_FS002V02(Without weighing)",                 # parcel feedline assembly
    "Static_Infeed Or Orientation Conveyor (FAL_F014V02)_01",  # infeed assembly
    "Static_Infeed Or Orientation Conveyor",          # seen once sometimes
]

# Feedline name used for feedline/residual math
FEEDLINE_BLOCK = "FAL_FS002V02(Without weighing)"


# =============================================================================
# Filters
# =============================================================================
ANON_RE = re.compile(r"^\*?[UD]\d+$", re.IGNORECASE)  # *U138, U138, *D12, D12


def is_noise_block(name: str) -> bool:
    if not name:
        return True
    n = name.strip()
    if not n:
        return True
    if n.startswith("*"):
        return True
    if ANON_RE.match(n):
        return True
    if n.upper().startswith("ACAD_") or n.upper().startswith("A$C"):
        return True
    return False


def is_excluded(name: str) -> bool:
    up = name.strip().upper()
    for p in EXCLUDE_PREFIXES:
        if up.startswith(p.strip().upper()):
            return True
    for ex in EXCLUDE_EXACT:
        if up == ex.strip().upper():
            return True
    return False


def keep_block(name: str) -> bool:
    if not name or not name.strip():
        return False
    if is_excluded(name):
        return False
    if (not INCLUDE_NOISE) and is_noise_block(name):
        return False
    return True


# =============================================================================
# Table printer
# =============================================================================
def print_table(title: str, rows, limit: int = 0) -> None:
    if limit and limit > 0:
        rows = rows[:limit]

    print(f"\n{title}")
    if not rows:
        print("(no rows)")
        return

    idx_w = max(1, len(str(len(rows))))
    name_w = max(9, max((len(r[0]) for r in rows), default=9))
    count_w = max(5, max((len(str(r[1])) for r in rows), default=5))

    sep = f"+-{'-'*idx_w}-+-{'-'*name_w}-+-{'-'*count_w}-+"
    print(sep)
    print(f"| {'#'.rjust(idx_w)} | {'Component'.ljust(name_w)} | {'Count'.rjust(count_w)} |")
    print(sep)

    for i, (name, cnt) in enumerate(rows, 1):
        print(f"| {str(i).rjust(idx_w)} | {name.ljust(name_w)} | {str(cnt).rjust(count_w)} |")

    print(sep)


def sort_rows(counter: Counter) -> list[tuple[str, int]]:
    rows = list(counter.items())
    if SORT_BY == "name":
        rows.sort(key=lambda x: x[0].lower())
    else:
        rows.sort(key=lambda x: (-x[1], x[0].lower()))
    return rows


# =============================================================================
# Exploder (nested block expansion)
# =============================================================================
def build_block_children(doc) -> dict[str, Counter]:
    """
    Returns: block_name -> Counter(child_block_name -> count)
    Only includes child INSERTs that pass keep_block().
    """
    children: dict[str, Counter] = {}

    for b in doc.blocks:
        bname = b.name
        if bname.upper() in ("*MODEL_SPACE", "*PAPER_SPACE"):
            continue

        c = Counter()
        for e in b:
            if (e.dxftype() or "").upper() != "INSERT":
                continue
            child = getattr(e.dxf, "name", None)
            if isinstance(child, str):
                child = child.strip()
                if keep_block(child):
                    c[child] += 1

        children[bname] = c

    return children


def scale_counter(c: Counter, factor: int) -> Counter:
    """Return a new Counter where each value is multiplied by factor."""
    if factor == 1:
        return c.copy()
    out = Counter()
    for k, v in c.items():
        out[k] = v * factor
    return out


def explode_block(
    block_name: str,
    block_children: dict[str, Counter],
    memo: dict[str, Counter],
    visiting: set[str],
) -> Counter:
    """
    Returns expanded composition of a block.

    If LEAF_ONLY=True:
      - If block has no non-noise child inserts => leaf => {block_name: 1}
      - Else returns sum of expanded children
    If LEAF_ONLY=False:
      - Includes container itself as {block_name: 1} + expanded children
    """
    if block_name in memo:
        return memo[block_name].copy()

    if block_name in visiting:
        # cycle guard
        return Counter()

    visiting.add(block_name)

    kids = block_children.get(block_name, Counter())
    if not kids:
        out = Counter({block_name: 1})
    else:
        out = Counter()
        if not LEAF_ONLY:
            out[block_name] += 1

        for child, cnt in kids.items():
            child_expanded = explode_block(child, block_children, memo, visiting)
            out += scale_counter(child_expanded, cnt)

    visiting.remove(block_name)
    memo[block_name] = out.copy()
    return out


# =============================================================================
# Helpers: block breakdown + parent finder
# =============================================================================
def show_block_breakdown(block_children: dict[str, Counter], block_name: str) -> None:
    if block_name not in block_children:
        print(f"\nBreakdown: {block_name} (not found in doc.blocks)")
        return
    kids = block_children.get(block_name, Counter())
    print(f"\nBreakdown: {block_name}")
    if not kids:
        print("  (no child INSERTs)")
        return
    for k, v in kids.most_common():
        print(f"  {k} = {v}")


def find_parents_of(block_children: dict[str, Counter], child_name: str) -> list[tuple[str, int]]:
    parents = []
    for parent, kids in block_children.items():
        c = kids.get(child_name, 0)
        if c > 0:
            parents.append((parent, c))
    parents.sort(key=lambda x: (-x[1], x[0].lower()))
    return parents


# =============================================================================
# Main
# =============================================================================
def main():
    dxf_path = Path(INPUT_DXF)
    if not dxf_path.exists():
        raise SystemExit(f"DXF not found: {dxf_path.resolve()}")

    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    # TABLE 1: modelspace INSERT counts (top-level)
    top_counts = Counter()
    top_insert_total = 0
    modelspace_inserts: list[str] = []

    for e in msp:
        if (e.dxftype() or "").upper() != "INSERT":
            continue

        top_insert_total += 1
        name = getattr(e.dxf, "name", None)
        if not isinstance(name, str):
            continue

        name = name.strip()
        if not keep_block(name):
            continue

        top_counts[name] += 1
        modelspace_inserts.append(name)

    # Build block children graph
    block_children = build_block_children(doc)

    # Print breakdown for key blocks
    for b in TARGET_BLOCKS:
        show_block_breakdown(block_children, b)

    # Auto-detect bag-induct parent(s) based on the 30-degree turn block, and print those parents + breakdown
    bag_child = "FAL_PMC9V01(1000mm_30_deg)"
    parents = find_parents_of(block_children, bag_child)
    print(f"\nParents containing '{bag_child}':")
    if not parents:
        print("  (none found)")
    else:
        for p, c in parents[:20]:
            print(f"  {p} -> {c}")
        top_parent = parents[0][0]
        show_block_breakdown(block_children, top_parent)

    # TABLE 2: exploded counts (all modelspace inserts expanded)
    memo: dict[str, Counter] = {}
    exploded_counts = Counter()
    for name in modelspace_inserts:
        exploded_counts += explode_block(name, block_children, memo, set())

    # Feedline composition + residual analysis
    feedline_comp = explode_block(FEEDLINE_BLOCK, block_children, {}, set())
    feedline_count = top_counts.get(FEEDLINE_BLOCK, 0)

    feedline_total = Counter()
    for k, v in feedline_comp.items():
        feedline_total[k] = v * feedline_count

    residual = exploded_counts.copy()
    for k, v in feedline_total.items():
        residual[k] -= v
        if residual[k] <= 0:
            del residual[k]

    # Print summary + tables
    print(f"\nInput DXF: {dxf_path.name}")
    print(f"Total modelspace INSERT entities: {top_insert_total}")
    print(f"Unique components (top-level after filter): {len(top_counts)}")
    print(f"Unique components (exploded after filter): {len(exploded_counts)}")
    print(f"Explode mode: {'LEAF_ONLY' if LEAF_ONLY else 'INCLUDE_CONTAINERS'}")

    print_table("TABLE 1: Modelspace INSERT Counts", sort_rows(top_counts), TOP_N_TABLE1)
    print_table("TABLE 2: Exploded Component Counts (Nested INSERTs Expanded)", sort_rows(exploded_counts), TOP_N_TABLE2)

    print("\n====================")
    print("FEEDLINE (per 1) composition:")
    for k, v in feedline_comp.most_common():
        print(f"  {k} = {v}")

    print(f"\nFEEDLINE total contribution (x{feedline_count}):")
    for k, v in feedline_total.most_common():
        print(f"  {k} = {v}")

    print("\nRESIDUAL (exploded minus feedline contribution):")
    for k, v in sort_rows(residual):
        print(f"  {k} = {v}")
    print("====================\n")


if __name__ == "__main__":
    main()
