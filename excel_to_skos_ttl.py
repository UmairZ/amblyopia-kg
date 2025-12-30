#!/usr/bin/env python3
"""
Convert an Excel sheet to SKOS Turtle (.skos.ttl).

Requirements:
  pip install pandas openpyxl rdflib

Example:
  python excel_to_skos_ttl.py \
    --excel concepts.xlsx \
    --sheet Concepts \
    --base-uri "https://example.org/vocab/" \
    --scheme-id "my-scheme" \
    --scheme-label "My Concept Scheme" \
    --out my-scheme.skos.ttl
"""

from __future__ import annotations

import argparse
import sys
from typing import Optional, List

import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, SKOS, DCTERMS


def split_semicolon(value: Optional[str]) -> List[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    s = str(value).strip()
    if not s:
        return []
    return [part.strip() for part in s.split(";") if part.strip()]


def is_blank(value) -> bool:
    return value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == ""


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Excel to SKOS TTL.")
    parser.add_argument("--excel", required=True, help="Path to .xlsx file")
    parser.add_argument("--sheet", default="Concepts", help="Sheet name (default: Concepts)")
    parser.add_argument("--out", required=True, help="Output TTL filename (e.g., vocab.skos.ttl)")

    parser.add_argument("--base-uri", required=True, help="Base URI for concepts, e.g. https://example.org/vocab/")
    parser.add_argument("--scheme-id", required=True, help="ConceptScheme local id, e.g. my-scheme")
    parser.add_argument("--scheme-label", required=True, help="ConceptScheme label (human readable)")

    parser.add_argument("--lang", default="en", help="Language tag for labels/definitions (default: en)")
    parser.add_argument("--creator", default=None, help="Optional dcterms:creator literal")
    parser.add_argument("--issued", default=None, help="Optional dcterms:issued (ISO date string), e.g. 2025-12-16")

    args = parser.parse_args()

    base_uri = args.base_uri.rstrip("/") + "/"
    EX = Namespace(base_uri)

    scheme_uri = URIRef(base_uri + args.scheme_id)

    # ---- Read Excel
    try:
        df = pd.read_excel(args.excel, sheet_name=args.sheet, dtype=str)
    except Exception as e:
        print(f"Error reading Excel: {e}", file=sys.stderr)
        return 2

    # Normalize column names (strip spaces)
    df.columns = [str(c).strip() for c in df.columns]

    # Required columns
    required = ["id", "prefLabel"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"Missing required column(s): {missing}", file=sys.stderr)
        print(f"Found columns: {list(df.columns)}", file=sys.stderr)
        return 2

    # Optional columns we support
    col_id = "id"
    col_pref = "prefLabel"
    col_alt = "altLabel"        # semicolon separated
    col_def = "definition"
    col_broader = "broader"     # parent id
    col_notation = "notation"

    # ---- Build RDF
    g = Graph()
    g.bind("skos", SKOS)
    g.bind("dcterms", DCTERMS)
    g.bind("ex", EX)

    # ConceptScheme
    g.add((scheme_uri, RDF.type, SKOS.ConceptScheme))
    g.add((scheme_uri, SKOS.prefLabel, Literal(args.scheme_label, lang=args.lang)))

    if args.creator:
        g.add((scheme_uri, DCTERMS.creator, Literal(args.creator)))
    if args.issued:
        g.add((scheme_uri, DCTERMS.issued, Literal(args.issued)))

    # First pass: create concepts + basic props
    # We'll also keep a map from id -> concept URI for linking.
    concept_uri_by_id = {}

    for _, row in df.iterrows():
        cid = None if is_blank(row.get(col_id)) else str(row.get(col_id)).strip()
        pref = None if is_blank(row.get(col_pref)) else str(row.get(col_pref)).strip()

        if not cid or not pref:
            # Skip incomplete rows
            continue

        c_uri = URIRef(base_uri + cid)
        concept_uri_by_id[cid] = c_uri

        g.add((c_uri, RDF.type, SKOS.Concept))
        g.add((c_uri, SKOS.inScheme, scheme_uri))
        g.add((c_uri, SKOS.prefLabel, Literal(pref, lang=args.lang)))

        # alt labels
        if col_alt in df.columns:
            for alt in split_semicolon(row.get(col_alt)):
                g.add((c_uri, SKOS.altLabel, Literal(alt, lang=args.lang)))

        # definition
        if col_def in df.columns and not is_blank(row.get(col_def)):
            g.add((c_uri, SKOS.definition, Literal(str(row.get(col_def)).strip(), lang=args.lang)))

        # notation
        if col_notation in df.columns and not is_blank(row.get(col_notation)):
            g.add((c_uri, SKOS.notation, Literal(str(row.get(col_notation)).strip())))

        # Top concepts can be inferred later (if no broader), or you can add a column.
        # We'll infer in second pass.

    # Second pass: broader/narrower links + top concepts
    for _, row in df.iterrows():
        cid = None if is_blank(row.get(col_id)) else str(row.get(col_id)).strip()
        if not cid or cid not in concept_uri_by_id:
            continue

        c_uri = concept_uri_by_id[cid]

        parent_id = None
        if col_broader in df.columns and not is_blank(row.get(col_broader)):
            parent_id = str(row.get(col_broader)).strip()

        if parent_id:
            p_uri = concept_uri_by_id.get(parent_id)
            if p_uri is None:
                # If parent not found, you can either skip or mint it.
                # Here we warn and skip.
                print(f"Warning: broader '{parent_id}' referenced by '{cid}' not found in sheet; skipping link.",
                      file=sys.stderr)
            else:
                g.add((c_uri, SKOS.broader, p_uri))
                g.add((p_uri, SKOS.narrower, c_uri))
        else:
            # no broader => top concept
            g.add((scheme_uri, SKOS.hasTopConcept, c_uri))
            g.add((c_uri, SKOS.topConceptOf, scheme_uri))

    # ---- Serialize
    try:
        g.serialize(destination=args.out, format="turtle")
    except Exception as e:
        print(f"Error writing TTL: {e}", file=sys.stderr)
        return 2

    print(f"Wrote {args.out} with {len(concept_uri_by_id)} concepts.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())