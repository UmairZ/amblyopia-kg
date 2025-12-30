import pandas as pd
from pyvis.network import Network

# ---------- CONFIG ----------
XLSX_PATH = "concepts.xlsx"   # <-- change
SHEET_NAME = "Sheet1"            # <-- change (e.g., "Sheet2" if needed)
OUT_HTML = "amblyopia_kg.html"

# Your column names (exactly as they appear in Excel)
COL_ID = "CONCEPT ID"
COL_NAME = "PREFERRED NAME"
COL_CAT = "CONCEPT CATEGORY"
COL_PARENT = "PARENT"

# ---------- LOAD ----------
df = pd.read_excel(XLSX_PATH, sheet_name=SHEET_NAME, dtype=str)
df = df.fillna("")

for col in [COL_ID, COL_NAME, COL_CAT]:
    if col not in df.columns:
        raise ValueError(f"Missing required column: '{col}'. Found: {list(df.columns)}")

# Keep only needed columns if present
keep_cols = [c for c in [COL_ID, COL_NAME, COL_CAT, COL_PARENT] if c in df.columns]
df = df[keep_cols].copy()

# ---------- DEDUPE (same concept may appear in multiple note sections) ----------
def first_nonempty(series):
    for v in series:
        v = str(v).strip()
        if v:
            return v
    return ""

agg = {COL_NAME: first_nonempty, COL_CAT: first_nonempty}
if COL_PARENT in df.columns:
    agg[COL_PARENT] = first_nonempty

df = df.groupby(COL_ID, as_index=False).agg(agg)

# Build helper maps
id_to_name = {r[COL_ID].strip(): r[COL_NAME].strip() for _, r in df.iterrows()}
name_to_id = {v: k for k, v in id_to_name.items() if v}

# ---------- BUILD GRAPH ----------
net = Network(height="750px", width="100%", directed=True, bgcolor="#ffffff")
net.force_atlas_2based()

# --- colors ---
TYPE_COLOR = "#FFD966"      # light yellow for type nodes
CONCEPT_COLOR = "#9DC3E6"    # light blue for concept nodes
UNRESOLVED_COLOR = "#F8CBAD" # light red/orange for unresolved parent placeholders

# Add category/type nodes
categories = sorted(df[COL_CAT].dropna().unique())
for cat in categories:
    cat_id = f"TYPE::{cat}"
    net.add_node(
        cat_id,
        label=cat,
        shape="box",
        color=TYPE_COLOR
    )

# Add concept nodes
for _, row in df.iterrows():
    cid = str(row[COL_ID]).strip()
    name = str(row[COL_NAME]).strip() or cid
    cat = str(row[COL_CAT]).strip() or "Unknown"

    net.add_node(
        cid,
        label=name,
        title=f"{cid}\nType: {cat}",
        shape="ellipse",
        color=CONCEPT_COLOR
    )

# Add edges (IS_A first, track who has a resolved parent)
has_resolved_parent = set()

for _, row in df.iterrows():
    cid = str(row[COL_ID]).strip()

    parent_name = ""
    if COL_PARENT in df.columns:
        parent_name = str(row.get(COL_PARENT, "")).strip()

    if parent_name:
        parent_id = name_to_id.get(parent_name)
        if parent_id:
            # child -> parent
            net.add_edge(cid, parent_id, label="IS_A", arrows="to")
            has_resolved_parent.add(cid)
        else:
            # Parent referenced but not present as a concept row
            placeholder_id = f"UNRESOLVED::{parent_name}"
            if placeholder_id not in net.node_ids:
                net.add_node(
                    placeholder_id,
                    label=parent_name,
                    shape="ellipse",
                    color=UNRESOLVED_COLOR,
                    title="Parent referenced but not found as a concept row"
                )
            net.add_edge(cid, placeholder_id, label="IS_A", arrows="to")
            # parent unresolved -> do NOT mark as resolved parent

# Add HAS_TYPE edges only for roots or unresolved-parent concepts
for _, row in df.iterrows():
    cid = str(row[COL_ID]).strip()
    cat = str(row[COL_CAT]).strip() or "Unknown"

    # If this concept has a resolved parent, skip HAS_TYPE (inherit it)
    if cid in has_resolved_parent:
        continue

    net.add_edge(cid, f"TYPE::{cat}", label="HAS_TYPE", arrows="to")

net.write_html(OUT_HTML, open_browser=False, notebook=False)
print(f"✅ Wrote interactive KG visualization to: {OUT_HTML}")