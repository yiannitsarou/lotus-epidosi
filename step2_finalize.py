# -*- coding: utf-8 -*-
"""
Βήμα 2 — Finalize & Exports
- Περιλαμβάνει:
  • finalize_step2_assignments / lock_step2_results
  • export_step2_minimal_nextcol (παλιό "ελαφρύ")
  • export_step2_nextcol_full (ΝΕΟ, DEFAULT) → κρατά ΟΛΕΣ τις αρχικές στήλες
    και προσθέτει τη ΒΗΜΑ2_ΣΕΝΑΡΙΟ_N αμέσως δεξιά από τη ΒΗΜΑ1_ΣΕΝΑΡΙΟ_N,
    ένα φύλλο ανά σενάριο.
"""
from typing import Optional, Tuple, List, Dict
import pandas as pd
import re, math

# ------------------ Κλείδωμα Βήματος 2 ------------------
def finalize_step2_assignments(
    df: pd.DataFrame, 
    step2_col: str,
    final_col_name: Optional[str] = None
) -> Tuple[pd.DataFrame, dict]:
    if final_col_name is None:
        match = re.search(r'ΣΕΝΑΡΙΟ[_\s]*(\d+)', str(step2_col))
        scenario_id = match.group(1) if match else "1"
        final_col_name = f"ΤΕΛΙΚΟ_ΤΜΗΜΑ_ΣΕΝΑΡΙΟ_{scenario_id}"
    result_df = df.copy()
    result_df[final_col_name] = result_df[step2_col].copy()
    unplaced_mask = pd.isna(result_df[final_col_name])
    unplaced_count = unplaced_mask.sum()
    if unplaced_count == 0:
        stats = {
            "total_students": len(result_df),
            "already_placed": len(result_df),
            "newly_placed": 0,
            "class_distribution": result_df[final_col_name].value_counts().to_dict()
        }
        return result_df, stats
    placed_classes = result_df[~unplaced_mask][final_col_name].value_counts()
    available_classes = sorted(placed_classes.index.tolist())
    if not available_classes:
        num_classes = max(2, math.ceil(len(result_df) / 25))
        available_classes = [f"Α{i+1}" for i in range(num_classes)]
        placed_classes = pd.Series([0] * len(available_classes), index=available_classes)
    unplaced_names = result_df[unplaced_mask]["ΟΝΟΜΑ"].tolist()
    classes_by_size = placed_classes.sort_values().index.tolist()
    for i, student_name in enumerate(unplaced_names):
        target_class = classes_by_size[i % len(classes_by_size)]
        student_idx = result_df[result_df["ΟΝΟΜΑ"] == student_name].index[0]
        result_df.loc[student_idx, final_col_name] = target_class
    final_distribution = result_df[final_col_name].value_counts().to_dict()
    stats = {
        "total_students": len(result_df),
        "already_placed": len(result_df) - unplaced_count,
        "newly_placed": unplaced_count,
        "class_distribution": final_distribution,
        "min_class_size": min(final_distribution.values()) if final_distribution else 0,
        "max_class_size": max(final_distribution.values()) if final_distribution else 0
    }
    return result_df, stats

def validate_final_assignments(df: pd.DataFrame, final_col: str) -> dict:
    validation = {
        "total_students": len(df),
        "students_with_assignment": (~pd.isna(df[final_col])).sum(),
        "students_without_assignment": pd.isna(df[final_col]).sum(),
        "is_complete": pd.isna(df[final_col]).sum() == 0,
        "unique_classes": df[final_col].nunique(),
        "class_list": sorted(df[final_col].dropna().unique().tolist())
    }
    if validation["is_complete"]:
        class_sizes = df[final_col].value_counts()
        validation.update({
            "min_class_size": class_sizes.min(),
            "max_class_size": class_sizes.max(),
            "avg_class_size": class_sizes.mean(),
            "class_size_std": class_sizes.std()
        })
    return validation

def lock_step2_results(df: pd.DataFrame, step2_column: str) -> pd.DataFrame:
    final_df, stats = finalize_step2_assignments(df, step2_column)
    print("=== ΚΛΕΙΔΩΜΑ ΒΗΜΑΤΟΣ 2 ===")
    print(f"Συνολικά παιδιά: {stats['total_students']}")
    print(f"Ήδη τοποθετημένα: {stats['already_placed']}")
    print(f"Νέες τοποθετήσεις: {stats['newly_placed']}")
    print(f"Κατανομή ανά τμήμα:")
    for class_name, count in sorted(stats['class_distribution'].items()):
        print(f"  {class_name}: {count} παιδιά")
    return final_df

# ------------------ Exporters ------------------
def export_step2_minimal_nextcol(
    step1_workbook_path: str,
    out_xlsx_path: str,
    *,
    seed: int = 42,
    max_results: int = 5,
    core_columns: Optional[List[str]] = None,
    sheet_naming: str = "ΣΕΝΑΡΙΟ_{id}"
) -> None:
    """Παλιός ελαφρύς exporter: κρατά βασικές στήλες + ΒΗΜΑ1/ΒΗΜΑ2."""
    from step_2_zoiroi_idiaterotites_FIXED_v3_PATCHED import step2_apply_FIXED_v3
    from step_2_helpers_FIXED import (
        normalize_columns, extract_step1_id, find_step1_scenario_columns, pick_core_columns
    )

    xls = pd.ExcelFile(step1_workbook_path)
    seen_ids = set()
    outputs: Dict[int, Dict] = {}

    for sh in xls.sheet_names:
        df_raw = xls.parse(sh)
        df = normalize_columns(df_raw)
        step1_cols = find_step1_scenario_columns(df)
        for step1_col in step1_cols:
            sid = extract_step1_id(step1_col)
            if sid in seen_ids:
                continue
            seen_ids.add(sid)

            options = step2_apply_FIXED_v3(df, step1_col, seed=seed, max_results=max_results)
            def key_fn(opt):
                label, opt_df, m = opt
                pen = m.get("penalty") if m.get("penalty") is not None else 10**9
                bro = m.get("broken") if m.get("broken") is not None else 10**9
                ped = m.get("ped_conflicts") if m.get("ped_conflicts") is not None else 10**9
                return (pen, bro, ped)
            best_label, best_df, best_metrics = sorted(options, key=key_fn)[0]

            step2_col = f"ΒΗΜΑ2_ΣΕΝΑΡΙΟ_{sid}"
            if step2_col not in best_df.columns:
                cands = [c for c in best_df.columns if str(c).startswith("ΒΗΜΑ2_")]
                if not cands:
                    raise RuntimeError(f"Δεν βρέθηκε στήλη ΒΗΜΑ2 για το σενάριο {sid}.")
                step2_col = cands[0]

            keep_core = pick_core_columns(best_df, core_columns)
            cols = keep_core + [step1_col, step2_col]
            minimal_df = best_df[cols].copy()

            outputs[sid] = {"sheet_name": sheet_naming.format(id=sid), "df": minimal_df}

    with pd.ExcelWriter(out_xlsx_path, engine="xlsxwriter") as writer:
        for sid in sorted(outputs.keys()):
            outputs[sid]["df"].to_excel(writer, sheet_name=outputs[sid]["sheet_name"], index=False)

def export_step2_nextcol_full(
    step1_workbook_path: str,
    out_xlsx_path: str,
    *,
    seed: int = 42,
    max_results: int = 5,
    sheet_naming: str = "ΣΕΝΑΡΙΟ_{id}"
) -> None:
    """
    ΝΕΟΣ DEFAULT EXPORTER — FULL:
    - Κρατά ΟΛΕΣ τις αρχικές στήλες του workbook του Βήματος 1 (χωρίς καμία αφαίρεση).
    - Εκτελεί Βήμα 2 ανά σενάριο και προσθέτει τη «ΒΗΜΑ2_ΣΕΝΑΡΙΟ_{N}»
      αμέσως δεξιά από τη «ΒΗΜΑ1_ΣΕΝΑΡΙΟ_{N}». Ένα sheet ανά σενάριο.
    - Δεν γράφει καμία FINAL/audit στήλη.
    """
    from step_2_zoiroi_idiaterotites_FIXED_v3_PATCHED import step2_apply_FIXED_v3

    xls = pd.ExcelFile(step1_workbook_path)
    used_ids = set()
    outputs: Dict[int, Dict] = {}

    def _sid_from_col(col_name: str) -> int:
        m = re.search(r'ΣΕΝΑΡΙΟ[_\s]*(\d+)', str(col_name).upper())
        return int(m.group(1)) if m else 1

    def _find_step1_cols(df: pd.DataFrame):
        return [c for c in df.columns if str(c).strip().upper().startswith("ΒΗΜΑ1_ΣΕΝΑΡΙΟ_")]

    for sh in xls.sheet_names:
        orig_df = xls.parse(sh)
        step1_cols = _find_step1_cols(orig_df)
        for step1_col in step1_cols:
            sid = _sid_from_col(step1_col)
            if sid in used_ids:
                continue
            used_ids.add(sid)

            options = step2_apply_FIXED_v3(orig_df.copy(), step1_col, seed=seed, max_results=max_results)
            def key_fn(opt):
                label, opt_df, m = opt
                pen = m.get("penalty") if m.get("penalty") is not None else 10**9
                bro = m.get("broken") if m.get("broken") is not None else 10**9
                ped = m.get("ped_conflicts") if m.get("ped_conflicts") is not None else 10**9
                return (pen, bro, ped)
            best_label, best_df, best_metrics = sorted(options, key=key_fn)[0]

            step2_col = f"ΒΗΜΑ2_ΣΕΝΑΡΙΟ_{sid}"
            if step2_col not in best_df.columns:
                cands = [c for c in best_df.columns if str(c).startswith("ΒΗΜΑ2_")]
                if not cands:
                    raise RuntimeError(f"Δεν βρέθηκε στήλη ΒΗΜΑ2 στο αποτέλεσμα για σενάριο {sid}.")
                step2_col = cands[0]

            if "ΟΝΟΜΑ" not in orig_df.columns:
                raise RuntimeError("Το αρχικό φύλλο δεν έχει στήλη 'ΟΝΟΜΑ'.")
            s_step2 = best_df.set_index("ΟΝΟΜΑ")[step2_col]
            merged = orig_df.copy()
            merged[step2_col] = merged["ΟΝΟΜΑ"].map(s_step2.to_dict())

            cols = merged.columns.tolist()
            if step2_col in cols:
                cols.remove(step2_col)
            idx = cols.index(step1_col) + 1 if step1_col in cols else len(cols)
            cols = cols[:idx] + [step2_col] + cols[idx:]
            merged = merged[cols]

            outputs[sid] = {"sheet_name": sheet_naming.format(id=sid), "df": merged}

    with pd.ExcelWriter(out_xlsx_path, engine="xlsxwriter") as writer:
        for sid in sorted(outputs.keys()):
            outputs[sid]["df"].to_excel(writer, sheet_name=outputs[sid]["sheet_name"], index=False)
