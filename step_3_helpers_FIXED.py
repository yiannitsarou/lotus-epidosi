
# -*- coding: utf-8 -*-
"""
step_3_helpers_FIXED.py
- ΦΙΛΟΙ parsing από string ή list
- Έλεγχος ΑΜΟΙΒΑΙΑΣ φιλίας (μόνο ΔΥΑΔΕΣ)
- Μέτρηση «σπασμένων» φιλικών ΔΥΑΔΩΝ (χωρίς διπλομέτρηση)
- Penalty score για Βήμα 3
- Επιλογή σεναρίων βάσει θεωρίας
"""

from typing import List, Tuple, Dict, Set
import pandas as pd
import re, ast

SAFE_SEP = re.compile(r"[,\|\;/·\n]+")

def parse_friends_string(x) -> List[str]:
    if isinstance(x, list):
        return [str(s).strip() for s in x if str(s).strip()]
    if pd.isna(x):
        return []
    s = str(x).strip()
    if not s:
        return []
    # Προσπάθεια ως Python list: "['Α','Β']"
    try:
        val = ast.literal_eval(s)
        if isinstance(val, list):
            return [str(t).strip() for t in val if str(t).strip()]
    except Exception:
        pass
    # Διαφορετικά, split με ασφαλές regex
    parts = SAFE_SEP.split(s)
    return [p.strip() for p in parts if p.strip() and p.strip().lower()!="nan"]

def are_mutual_pair(df: pd.DataFrame, a: str, b: str) -> bool:
    ra = df[df["ΟΝΟΜΑ"].astype(str)==str(a)]
    rb = df[df["ΟΝΟΜΑ"].astype(str)==str(b)]
    if ra.empty or rb.empty:
        return False
    fa = set(parse_friends_string(ra.iloc[0].get("ΦΙΛΟΙ","")))
    fb = set(parse_friends_string(rb.iloc[0].get("ΦΙΛΟΙ","")))
    return (str(b).strip() in fa) and (str(a).strip() in fb)

def mutual_dyads(df: pd.DataFrame) -> Set[Tuple[str,str]]:
    names = df["ΟΝΟΜΑ"].astype(str).str.strip().tolist()
    pairs: Set[Tuple[str,str]] = set()
    for i, a in enumerate(names):
        for b in names[i+1:]:
            if are_mutual_pair(df, a, b):
                pairs.add(tuple(sorted([a,b])))
    return pairs

def count_broken_dyads(before_df: pd.DataFrame, after_df: pd.DataFrame, scenario_col: str) -> int:
    """Μετρά πόσες αμοιβαίες ΔΥΑΔΕΣ σπάνε στο after_df (δηλ. κατανέμονται σε διαφορετικές τάξεις)."""
    pairs = mutual_dyads(before_df)
    name2class = {str(r["ΟΝΟΜΑ"]).strip(): str(r.get(scenario_col)) for _, r in after_df.iterrows() if pd.notna(r.get(scenario_col))}
    broken=0
    for a,b in pairs:
        ca = name2class.get(a); cb = name2class.get(b)
        if ca is None or cb is None:
            # αν κάποιος δεν έχει τοποθετηθεί, θεωρούμε ότι η δυάδα δεν διατηρήθηκε
            broken += 1
        elif ca != cb:
            broken += 1
    return broken

def calculate_penalty_score_step3(df: pd.DataFrame, scenario_col: str, num_classes: int) -> int:
    """+1 για κάθε μονάδα διαφοράς >2 σε αγόρια, κορίτσια, πληθυσμό."""
    penalty = 0
    boys_counts=[]; girls_counts=[]; pop_counts=[]
    for i in range(num_classes):
        cl = f"Α{i+1}"
        sub = df[df[scenario_col]==cl]
        boys_counts.append(int((sub["ΦΥΛΟ"].astype(str).str.upper()=="Α").sum()))
        girls_counts.append(int((sub["ΦΥΛΟ"].astype(str).str.upper()=="Κ").sum()))
        pop_counts.append(len(sub))
    if boys_counts:
        penalty += max(0, max(boys_counts)-min(boys_counts)-2)
    if girls_counts:
        penalty += max(0, max(girls_counts)-min(girls_counts)-2)
    if pop_counts:
        penalty += max(0, max(pop_counts)-min(pop_counts)-2)
    return int(penalty)

def select_best_scenarios(results: List[Tuple[str, pd.DataFrame, Dict]]) -> List[Tuple[str,pd.DataFrame,Dict]]:
    """
    results: [(sheet_name, df_after, meta), ...], όπου meta περιέχει:
       {"broken": int, "penalty": int}
    Κανόνες:
      - Αν υπάρχουν σενάρια με broken==0 → επέλεξε όσα έχουν το μικρότερο penalty (έως 5)
      - Αλλιώς → επέλεξε όσα έχουν το μικρότερο broken, και tie-break με penalty (έως 5)
    """
    if not results:
        return []
    zero = [t for t in results if t[2].get("broken", 0)==0]
    if zero:
        zero.sort(key=lambda x: x[2].get("penalty", 0))
        return zero[:5]
    # αλλιώς
    results.sort(key=lambda x: (x[2].get("broken", 1_000_000), x[2].get("penalty", 1_000_000)))
    return results[:5]
