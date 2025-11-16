# ---------- DEBUGGING HELPERS (temporary, remove after debug) ----------
import json, sys

def _norm_for_debug(s):
    if not isinstance(s, str):
        return ""
    return " ".join(s.upper().replace("SUBTOTAL:", "").split())

print("DEBUG: total rows in df =", len(df))
# show first 10 rows for quick sanity
print("DEBUG: sample rows (first 10):")
print(df.head(10).to_string(index=False))

# Build list of sections (same as code)
subtotal_mask = df["Legal Entity"].str.upper().str.contains("SUBTOTAL", na=False)
subtotal_indices = list(df.index[subtotal_mask])
sections_debug = []
start = 0
for idx in subtotal_indices:
    section = df.iloc[start: idx+1].copy()
    sections_debug.append(section)
    start = idx + 1
if start < len(df):
    sections_debug.append(df.iloc[start:].copy())

print(f"DEBUG: detected {len(sections_debug)} sections (subtotal count={len(subtotal_indices)})")

# Build section_totals_debug using subtotal rows (as current code does)
section_totals_debug = {}
for i, section in enumerate(sections_debug):
    sec = section.reset_index(drop=True)
    if len(sec) == 0:
        print(f"DEBUG: section {i} empty")
        continue
    subtotal_row = sec.iloc[-1]
    legal_norm = _norm_for_debug(subtotal_row["Legal Entity"])
    try:
        gs_value = float(subtotal_row.get("GS Commitment", 0) or 0)
    except Exception:
        gs_value = 0
    section_totals_debug[legal_norm] = gs_value
    print(f"DEBUG: section {i} subtotal_legal='{subtotal_row['Legal Entity']}' -> legal_norm='{legal_norm}' gs={gs_value} rows={len(sec)}")

# Now show each section's FEEDER attempts and matches
for i, section in enumerate(sections_debug):
    sec = section.reset_index(drop=True)
    if len(sec) == 0:
        continue
    subtotal_legal_norm = _norm_for_debug(sec.iloc[-1]["Legal Entity"])
    print(f"\nDEBUG: SECTION {i} subtotal_legal_norm = '{subtotal_legal_norm}'")
    data_rows = sec.iloc[:len(sec)-1]
    for r_idx, row in data_rows.reset_index(drop=True).iterrows():
        bin_id = str(row.get("Bin ID",""))
        is_feeder = "FEEDER" in bin_id.upper()
        if is_feeder:
            feeder_legal = str(row.get("Legal Entity","")).upper().replace("HOLDING","")
            feeder_legal = " ".join(feeder_legal.split())
            print(f"  FEEDER row pos={r_idx} BinID='{bin_id}' feeder_legal='{feeder_legal}'")
            matched_gs = section_totals_debug.get(feeder_legal)
            print(f"    -> lookup in section_totals_debug yields: {matched_gs} (exists={feeder_legal in section_totals_debug})")
