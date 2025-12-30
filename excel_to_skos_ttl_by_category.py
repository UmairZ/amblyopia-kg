#!/usr/bin/env python3
"""
Convert an Excel concept list into one SKOS Turtle file per "Concept Category" section.

Input columns (must match exactly):
  CONCEPT CATEGORY, CONCEPT ID, PREFERRED NAME, NOTES, SYNONYMS, PARENT

Category header rows are identified by:
  CONCEPT CATEGORY == "Concept Category"
In those rows:
  - CONCEPT ID is used as the ConceptScheme ID (and as the root concept ID)
  - PREFERRED NAME is used as the ConceptScheme prefLabel (and root concept prefLabel)
  - NOTES is used as skos:definition (optional)

All following rows until the next header belong to that scheme.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, SKOS


REQUIRED_COLS = [
    "CONCEPT CATEGORY",
    "CONCEPT ID",
    "PREFERRED NAME",
    "NOTES",
    "SYNONYMS",
    "PARENT",
]


def is_blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def split_synonyms(v: Optional[str]) -> List[str]:
    if is_blank(v):
        return []
    # Your sheet looks semicolon-separated, but this supports commas too.
    s = str(v).strip()
    parts = re.split(r"\s*;\s*|\s*,\s*", s)
    return [p.strip() for p in parts if p.strip()]


def safe_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\-\.]+", "_", name)
    return name


@dataclass
class CategoryBlock:
    scheme_id: str            # from header row CONCEPT ID, e.g. ConceptsDiagnosis
    scheme_label: str         # from header row PREFERRED NAME
    scheme_definition: str    # from header row NOTES (optional)
    rows: List[dict]          # concept rows following header


def parse_blocks(df: pd.DataFrame) -> List[CategoryBlock]:
    blocks: List[CategoryBlock] = []
    current: Optional[CategoryBlock] = None

    for _, r in df.iterrows():
        cat = "" if is_blank(r["CONCEPT CATEGORY"]) else str(r["CONCEPT CATEGORY"]).strip()
        cid = "" if is_blank(r["CONCEPT ID"]) else str(r["CONCEPT ID"]).strip()
        label = "" if is_blank(r["PREFERRED NAME"]) else str(r["PREFERRED NAME"]).strip()
        notes = "" if is_blank(r["NOTES"]) else str(r["NOTES"]).strip()
        syn = "" if is_blank(r["SYNONYMS"]) else str(r["SYNONYMS"]).strip()
        parent = "" if is_blank(r["PARENT"]) else str(r["PARENT"]).strip()

        # Skip completely empty rows
        if not any([cat, cid, label, notes, syn, parent]):
            continue

        if cat == "Concept Category":
            # Close previous block
            if current is not None:
                blocks.append(current)

            if not cid or not label:
                raise ValueError(
                    f"Found a Concept Category header row missing CONCEPT ID or PREFERRED NAME: "
                    f"CONCEPT ID='{cid}', PREFERRED NAME='{label}'"
                )

            current = CategoryBlock(
                scheme_id=cid,
                scheme_label=label,
                scheme_definition=notes,
                rows=[],
            )
        else:
            # Normal concept row must belong to a block
            if current is None:
                # Ignore rows before first header, or raise if you prefer:
                # raise ValueError("Found concept rows before any Concept Category header row.")
                continue

            # Must have ID + label to become a concept
            if not cid or not label:
                continue

            current.rows.append(
                {
                    "concept_category": cat,  # e.g. Diagnosis, Symptom, Modifier...
                    "id": cid,
                    "prefLabel": label,
                    "definition": notes,
                    "synonyms": syn,
                    "parent": parent,
                }
            )

    if current is not None:
        blocks.append(current)

    return blocks


def build_graph_for_block(
    block: CategoryBlock,
    base_uri: str,
    lang: str,
) -> Tuple[Graph, str]:
    """
    Returns (graph, output_filename_stub)
    Also prints warnings for PARENT values that don't match any CONCEPT ID in this block.
    """
    base_uri = base_uri.rstrip("/") + "/"
    EX = Namespace(base_uri)

    scheme_uri = URIRef(base_uri + block.scheme_id)
    root_concept_uri = URIRef(base_uri + block.scheme_id)

    g = Graph()
    g.bind("skos", SKOS)
    g.bind("ex", EX)

    # ConceptScheme
    g.add((scheme_uri, RDF.type, SKOS.ConceptScheme))
    g.add((scheme_uri, SKOS.prefLabel, Literal(block.scheme_label, lang=lang)))
    if block.scheme_definition:
        g.add((scheme_uri, SKOS.definition, Literal(block.scheme_definition, lang=lang)))

    # Root concept (same ID as scheme)
    g.add((root_concept_uri, RDF.type, SKOS.Concept))
    g.add((root_concept_uri, SKOS.inScheme, scheme_uri))
    g.add((root_concept_uri, SKOS.prefLabel, Literal(block.scheme_label, lang=lang)))
    if block.scheme_definition:
        g.add((root_concept_uri, SKOS.definition, Literal(block.scheme_definition, lang=lang)))

    g.add((scheme_uri, SKOS.hasTopConcept, root_concept_uri))
    g.add((root_concept_uri, SKOS.topConceptOf, scheme_uri))

    # First pass: create all concept URIs (including ones referenced as parents)
    concept_uri_by_id: Dict[str, URIRef] = {block.scheme_id: root_concept_uri}
    known_ids: set[str] = {block.scheme_id}

    for row in block.rows:
        known_ids.add(row["id"])
        concept_uri_by_id.setdefault(row["id"], URIRef(base_uri + row["id"]))

    # Track parent values that are unknown (not IDs)
    unknown_parents: set[str] = set()

    for row in block.rows:
        p = row["parent"]
        if p:
            if p not in known_ids:
                unknown_parents.add(p)
            # still mint it so serialization doesn't crash; you will clean manually
            concept_uri_by_id.setdefault(p, URIRef(base_uri + p))

    if unknown_parents:
        print(
            f"WARNING [{block.scheme_id}]: PARENT values not found as CONCEPT ID (likely labels/typos): "
            f"{sorted(unknown_parents)}",
            file=sys.stderr,
        )

    # Second pass: add concept triples
    for row in block.rows:
        cid = row["id"]
        c_uri = concept_uri_by_id[cid]

        g.add((c_uri, RDF.type, SKOS.Concept))
        g.add((c_uri, SKOS.inScheme, scheme_uri))
        g.add((c_uri, SKOS.prefLabel, Literal(row["prefLabel"], lang=lang)))

        if row["definition"]:
            g.add((c_uri, SKOS.definition, Literal(row["definition"], lang=lang)))

        for alt in split_synonyms(row["synonyms"]):
            g.add((c_uri, SKOS.altLabel, Literal(alt, lang=lang)))

        parent_id = row["parent"] if row["parent"] else block.scheme_id
        p_uri = concept_uri_by_id[parent_id]

        if parent_id != cid:
            g.add((c_uri, SKOS.broader, p_uri))

    out_stub = safe_filename(block.scheme_id)
    return g, out_stub

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="Path to concepts .xlsx")
    ap.add_argument("--sheet", default="Concepts", help="Sheet name (default: Concepts)")
    ap.add_argument("--base-uri", required=True, help="Base URI, e.g. https://example.org/vocab/")
    ap.add_argument("--outdir", default="out_ttl", help="Output directory (default: out_ttl)")
    ap.add_argument("--lang", default="en", help="Language tag (default: en)")
    ap.add_argument("--all-out", default=None, help="Optional: also write a single combined TTL file (e.g. concepts.skos.ttl)")
    args = ap.parse_args()

    try:
        df = pd.read_excel(args.excel, sheet_name=args.sheet, dtype=str)
    except Exception as e:
        print(f"Error reading Excel: {e}", file=sys.stderr)
        return 2

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        print(f"Missing required column(s): {missing}", file=sys.stderr)
        print(f"Found columns: {list(df.columns)}", file=sys.stderr)
        return 2

    blocks = parse_blocks(df)

    if not blocks:
        print("No Concept Category blocks found. Make sure CONCEPT CATEGORY has 'Concept Category' rows.", file=sys.stderr)
        return 2

    # os.makedirs(args.outdir, exist_ok=True)

    # for block in blocks:
    #     g, stub = build_graph_for_block(block, args.base_uri, args.lang)
    #     out_path = os.path.join(args.outdir, f"{stub}.skos.ttl")
    #     g.serialize(destination=out_path, format="turtle")
    #     print(f"Wrote {out_path}  (concepts: {len(block.rows)})")

    # return 0

    os.makedirs(args.outdir, exist_ok=True)

    combined = Graph()
    combined.bind("skos", SKOS)
    combined.bind("ex", Namespace(args.base_uri.rstrip("/") + "/"))

    total_concepts = 0

    for block in blocks:
        g, stub = build_graph_for_block(block, args.base_uri, args.lang)
        out_path = os.path.join(args.outdir, f"{stub}.skos.ttl")
        g.serialize(destination=out_path, format="turtle")
        print(f"Wrote {out_path}  (concepts: {len(block.rows)})")

        # Merge into combined graph
        for triple in g:
            combined.add(triple)

        total_concepts += len(block.rows)

    # Optional: write combined TTL
    if args.all_out:
        combined_path = args.all_out
        # If they passed a filename without a directory, write it inside outdir
        if not os.path.isabs(combined_path) and os.path.dirname(combined_path) == "":
            combined_path = os.path.join(args.outdir, combined_path)

        combined.serialize(destination=combined_path, format="turtle")
        print(f"Wrote {combined_path}  (total concept rows: {total_concepts}, total triples: {len(combined)})")

if __name__ == "__main__":
    raise SystemExit(main())