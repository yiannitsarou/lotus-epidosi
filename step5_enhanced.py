# -*- coding: utf-8 -*-
"""
step_5_ypoloipoi_mathites_CORRECTED_ENHANCED.py

ΔΙΟΡΘΩΣΕΙΣ & ΒΕΛΤΙΩΣΕΙΣ:
1. Penalty weights σύμφωνα με τις οδηγίες (+1 για όλα εκτός από σπασμένη φιλία +5)
2. Ισορροπία φύλου υπολογίζεται σε όλα τα τμήματα (όχι μόνο στους candidates)
3. Αφαιρέθηκε το RANDOM_SEED για καλύτερη τυχαιότητα στην ισοβαθμία
4. Βελτιώθηκε η λογική επιλογής τμήματος
5. ΠΡΟΣΘΗΚΗ: Προτίμηση υποψηφίων που κρατούν διαφορά πληθυσμού ≤2
6. ΠΡΟΣΘΗΚΗ: Δυναμικός υπολογισμός σπασμένων φιλιών αν λείπει η στήλη
7. ΠΡΟΣΘΗΚΗ: Εμπλουτισμένη επιστροφή με penalty score και όνομα νικητή
"""

from __future__ import annotations
import random, re
from typing import List, Dict, Tuple, Any, Optional
import pandas as pd

def _auto_num_classes(df: pd.DataFrame, override: Optional[int] = None) -> int:
    """Αυτόματος υπολογισμός αριθμού τμημάτων (25 μαθητές/τμήμα, min=2)."""
    import math
    n = len(df)
    k = max(2, math.ceil(n/25))
    return int(k if override is None else override)

# Tokens για boolean parsing
YES_TOKENS = {"Ν", "ΝΑΙ", "YES", "Y", "TRUE", "1"}
NO_TOKENS  = {"Ο", "ΟΧΙ", "NO", "N", "FALSE", "0"}

def _norm_str(x: Any) -> str:
    """Κανονικοποίηση string."""
    return str(x).strip().upper()

def _is_yes(x: Any) -> bool:
    """Έλεγχος αν η τιμή είναι 'ναι'."""
    return _norm_str(x) in YES_TOKENS

def _is_no(x: Any) -> bool:
    """Έλεγχος αν η τιμή είναι 'όχι'."""
    return _norm_str(x) in NO_TOKENS

def _parse_list_cell(x: Any) -> List[str]:
    """Parsing λίστας από διάφορα formats (string, list, κ.ά.)."""
    if isinstance(x, list):
        return [str(t).strip() for t in x if str(t).strip()]
    
    s = "" if pd.isna(x) else str(x)
    s = s.strip()
    if not s or s.upper() == "NAN":
        return []
    
    # Δοκιμή python list parsing
    try:
        v = eval(s, {}, {})
        if isinstance(v, list):
            return [str(t).strip() for t in v if str(t).strip()] 
    except Exception:
        pass
    
    # Split με διάφορους διαχωριστές
    parts = re.split(r"[,\|\;/·\n]+", s)
    return [p.strip() for p in parts if p.strip()]

def _is_good_greek(row: pd.Series) -> bool:
    """Έλεγχος καλής γνώσης ελληνικών (backward/forward compatible)."""
    if "ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ" in row:
        return _is_yes(row.get("ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"))
    if "ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ" in row:
        return _norm_str(row.get("ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ")) in {"ΚΑΛΗ", "GOOD", "Ν"}
    return False

def _get_class_labels(df: pd.DataFrame, scenario_col: str) -> List[str]:
    """Επιστρέφει τα labels των τμημάτων (Α1, Α2, ...)."""
    labs = sorted([str(v) for v in df[scenario_col].dropna().unique() 
                   if re.match(r"^Α\d+$", str(v))])
    return labs or [f"Α{i+1}" for i in range(2)]

def _count_broken_pairs(df: pd.DataFrame, scenario_col: str) -> int:
    """Δυναμικός υπολογισμός σπασμένων πλήρως αμοιβαίων φιλιών."""
    by_class = dict(zip(df["ΟΝΟΜΑ"].astype(str).str.strip(), df[scenario_col].astype(str)))
    broken = set()
    
    for _, r in df.iterrows():
        if not _is_yes(r.get("ΠΛΗΡΩΣ_ΑΜΟΙΒΑΙΑ", False)):
            continue
            
        me = str(r["ΟΝΟΜΑ"]).strip()
        c_me = by_class.get(me)
        
        for fr in _parse_list_cell(r.get("ΦΙΛΟΙ", [])):
            if me < fr:  # Αποφυγή διπλής καταμέτρησης
                friend_row = df[df["ΟΝΟΜΑ"].astype(str).str.strip() == fr]
                if not friend_row.empty and _is_yes(friend_row.iloc[0].get("ΠΛΗΡΩΣ_ΑΜΟΙΒΑΙΑ", False)):
                    c_fr = by_class.get(fr)
                    if pd.notna(c_me) and pd.notna(c_fr) and c_me != c_fr:
                        broken.add((me, fr))
    
    return len(broken)

def calculate_penalty_score(df: pd.DataFrame, scenario_col: str, 
                          num_classes: Optional[int] = None) -> int:
    """
    Υπολογισμός penalty score σύμφωνα με τις οδηγίες:
    - Γνώση Ελληνικών: +1 για κάθε διαφορά > 2
    - Πληθυσμός: +1 για κάθε διαφορά > 1  
    - Φύλο: +1 για κάθε διαφορά > 1 (αγόρια ή κορίτσια)
    - Σπασμένη Φιλία: +5 για κάθε σπασμένη πλήρως αμοιβαία φιλία
    """
    labs = _get_class_labels(df, scenario_col)
    if num_classes is None:
        num_classes = _auto_num_classes(df, None)

    penalty = 0

    # 1. Ισορροπία Γνώσης Ελληνικών
    greek_counts = []
    for lab in labs:
        sub = df[df[scenario_col] == lab].copy()
        greek_counts.append(int(sub.apply(_is_good_greek, axis=1).sum()))
    
    if greek_counts:
        greek_diff = max(greek_counts) - min(greek_counts)
        penalty += max(0, greek_diff - 2)  # +1 για κάθε διαφορά > 2

    # 2. Ισορροπία Πληθυσμού  
    class_sizes = [int((df[scenario_col] == lab).sum()) for lab in labs]
    if class_sizes:
        pop_diff = max(class_sizes) - min(class_sizes)
        penalty += max(0, pop_diff - 1)  # +1 για κάθε διαφορά > 1

    # 3. Ισορροπία Φύλου
    boys_counts = [int(((df[scenario_col] == lab) & 
                       (df["ΦΥΛΟ"].astype(str).str.upper() == "Α")).sum()) 
                   for lab in labs]
    girls_counts = [int(((df[scenario_col] == lab) & 
                        (df["ΦΥΛΟ"].astype(str).str.upper() == "Κ")).sum()) 
                    for lab in labs]
    
    if boys_counts:
        boys_diff = max(boys_counts) - min(boys_counts)
        penalty += max(0, boys_diff - 1)  # +1 για κάθε διαφορά > 1
    
    if girls_counts:
        girls_diff = max(girls_counts) - min(girls_counts)
        penalty += max(0, girls_diff - 1)  # +1 για κάθε διαφορά > 1

    # 4. Σπασμένες Πλήρως Αμοιβαίες Φιλίες
    if "ΣΠΑΣΜΕΝΗ_ΦΙΛΙΑ" in df.columns:
        broken_friendships = int(df["ΣΠΑΣΜΕΝΗ_ΦΙΛΙΑ"].apply(_is_yes).sum())
    else:
        broken_friendships = _count_broken_pairs(df, scenario_col)
    
    penalty += 5 * broken_friendships  # +5 για κάθε σπασμένη φιλία

    return penalty

def step5_place_remaining_students(df: pd.DataFrame, scenario_col: str, 
                                 num_classes: Optional[int] = None) -> Tuple[pd.DataFrame, int]:
    """
    Βήμα 5: Τοποθέτηση υπολοίπων μαθητών χωρίς (πλήρως αμοιβαίες) φιλίες.
    
    Κριτήρια τοποθέτησης (με σειρά προτεραιότητας):
    1. Τμήμα με μικρότερο πληθυσμό (< 25 μαθητές)
    2. Σε ισοπαλία: προτίμηση όσων κρατούν διαφορά πληθυσμού ≤2
    3. Σε ισοπαλία: καλύτερη ισορροπία φύλου σε ΌΛΑ τα τμήματα
    """
    df = df.copy()
    labs = _get_class_labels(df, scenario_col)
    if num_classes is None:
        num_classes = _auto_num_classes(df, None)

    # Προετοιμασία δεδομένων
    friends_list = (df["ΦΙΛΟΙ"].map(_parse_list_cell) if "ΦΙΛΟΙ" in df.columns 
                   else pd.Series([[]]*len(df)))
    fully_mutual = (df["ΠΛΗΡΩΣ_ΑΜΟΙΒΑΙΑ"].apply(_is_yes) if "ΠΛΗΡΩΣ_ΑΜΟΙΒΑΙΑ" in df.columns 
                   else pd.Series([False]*len(df)))
    broken_friendship = (df["ΣΠΑΣΜΕΝΗ_ΦΙΛΙΑ"].apply(_is_yes) if "ΣΠΑΣΜΕΝΗ_ΦΙΛΙΑ" in df.columns 
                        else pd.Series([False]*len(df)))

    # Mask για μαθητές που χρειάζονται τοποθέτηση στο Βήμα 5
    mask_step5 = (
        df[scenario_col].isna() & 
        ((friends_list.map(len) == 0) |  # Χωρίς φίλους
         (~fully_mutual) |               # Όχι πλήρως αμοιβαίες φιλίες  
         (broken_friendship))            # Σπασμένες φιλίες
    )

    remaining_students = df[mask_step5].copy()

    # Διαδοχική τοποθέτηση κάθε μαθητή
    for _, row in remaining_students.iterrows():
        name = str(row["ΟΝΟΜΑ"]).strip()
        gender = str(row["ΦΥΛΟ"]).strip().upper()

        # 1. Εύρεση διαθέσιμων τμημάτων με ελάχιστο πληθυσμό
        class_sizes = {lab: int((df[scenario_col] == lab).sum()) for lab in labs}
        min_size = min(class_sizes.values())
        available_classes = [lab for lab, size in class_sizes.items() 
                           if size == min_size and size < 25]
        
        if not available_classes:
            continue  # Όλα τα τμήματα γεμάτα

        if len(available_classes) == 1:
            chosen_class = available_classes[0]
        else:
            # 2. Προτίμηση υποψηφίων που κρατούν διαφορά πληθυσμού ≤2
            candidates_with_pop_diff = []
            for candidate in available_classes:
                new_sizes = {lab: int((df[scenario_col] == lab).sum()) + (1 if lab == candidate else 0)
                           for lab in labs}
                pop_diff = max(new_sizes.values()) - min(new_sizes.values())
                candidates_with_pop_diff.append((candidate, pop_diff))
            
            # Φιλτράρισμα: προτίμηση όσων κρατούν pop_diff ≤ 2
            preferred_pool = [c for c, d in candidates_with_pop_diff if d <= 2]
            pool = preferred_pool if preferred_pool else [c for c, _ in candidates_with_pop_diff]
            
            if len(pool) == 1:
                chosen_class = pool[0]
            else:
                # 3. Ισορροπία φύλου - υπολογισμός για ΌΛΑ τα τμήματα
                best_score = float('inf')
                best_classes = []
                
                for candidate in pool:
                    # Προσομοίωση προσθήκης μαθητή στο candidate τμήμα
                    boys_counts = []
                    girls_counts = []
                    
                    for lab in labs:
                        boys_in_class = int(((df[scenario_col] == lab) & 
                                           (df["ΦΥΛΟ"].astype(str).str.upper() == "Α")).sum())
                        girls_in_class = int(((df[scenario_col] == lab) & 
                                            (df["ΦΥΛΟ"].astype(str).str.upper() == "Κ")).sum())
                        
                        # Προσθήκη μαθητή στο candidate
                        if lab == candidate:
                            if gender == "Α":
                                boys_in_class += 1
                            elif gender == "Κ":
                                girls_in_class += 1
                        
                        boys_counts.append(boys_in_class)
                        girls_counts.append(girls_in_class)
                    
                    # Υπολογισμός διαφοράς φύλου
                    boys_diff = max(boys_counts) - min(boys_counts) if boys_counts else 0
                    girls_diff = max(girls_counts) - min(girls_counts) if girls_counts else 0
                    total_gender_diff = boys_diff + girls_diff
                    
                    if total_gender_diff < best_score:
                        best_score = total_gender_diff
                        best_classes = [candidate]
                    elif total_gender_diff == best_score:
                        best_classes.append(candidate)
                
                # Τυχαία επιλογή σε ισοπαλία
                chosen_class = random.choice(best_classes)

        # Τοποθέτηση μαθητή
        df.loc[df["ΟΝΟΜΑ"] == name, scenario_col] = chosen_class

    return df, calculate_penalty_score(df, scenario_col, num_classes)

def apply_step5_to_all_scenarios(scenarios_dict: Dict[str, pd.DataFrame], 
                               scenario_col: str, num_classes: Optional[int] = None) -> Tuple[pd.DataFrame, int, str]:
    """
    Εφαρμογή Βήματος 5 σε όλα τα σενάρια και επιλογή του βέλτιστου.
    
    Returns:
        Tuple[pd.DataFrame, int, str]: Το σενάριο με το χαμηλότερο penalty score, 
                                      το penalty score και το όνομα του σεναρίου
    """
    if not scenarios_dict:
        raise ValueError("Δεν δόθηκαν σενάρια προς επεξεργασία")
    
    results = {}
    for scenario_name, scenario_df in scenarios_dict.items():
        try:
            updated_df, score = step5_place_remaining_students(
                scenario_df, scenario_col, num_classes)
            results[scenario_name] = {"df": updated_df, "penalty_score": score}
        except Exception as e:
            print(f"Σφάλμα στο σενάριο {scenario_name}: {e}")
            continue

    if not results:
        raise ValueError("Κανένα σενάριο δεν επεξεργάστηκε επιτυχώς")

    # Εύρεση βέλτιστου σεναρίου
    min_score = min(v["penalty_score"] for v in results.values())
    best_scenarios = [k for k, v in results.items() if v["penalty_score"] == min_score]
    
    # Τυχαία επιλογή σε ισοβαθμία
    chosen_scenario = random.choice(best_scenarios)
    
    print(f"Επιλέχθηκε σενάριο: {chosen_scenario} με penalty score: {min_score}")
    return results[chosen_scenario]["df"], results[chosen_scenario]["penalty_score"], chosen_scenario


# Compatibility aliases για backward compatibility
step5_filikoi_omades = step5_place_remaining_students

def export_step5_like_template(step34_xlsx_path: str, out_xlsx_path: str) -> str:
    """
    Διαβάζει workbook με φύλλα 'ΣΕΝΑΡΙΟ_k' (12-στήλες template) και
    προσθέτει 'ΒΗΜΑ5_ΣΕΝΑΡΙΟ_k' εφαρμόζοντας το Βήμα 5 πάνω στο 'ΒΗΜΑ4_ΣΕΝΑΡΙΟ_k'.
    """
    xls = pd.ExcelFile(step34_xlsx_path)
    with pd.ExcelWriter(out_xlsx_path, engine="xlsxwriter") as writer:
        for sh in xls.sheet_names:
            if not str(sh).startswith("ΣΕΝΑΡΙΟ_"):
                continue
            df = xls.parse(sh)
            m = re.search(r"(\d+)$", sh)
            sid = int(m.group(1)) if m else 1
            col4 = f"ΒΗΜΑ4_ΣΕΝΑΡΙΟ_{sid}"
            col5 = f"ΒΗΜΑ5_ΣΕΝΑΡΙΟ_{sid}"
            if col4 not in df.columns:
                df.to_excel(writer, index=False, sheet_name=sh)
                continue
            updated_df, score = step5_place_remaining_students(df.copy(), scenario_col=col4, num_classes=None)
            base = ['Α/Α','ΟΝΟΜΑ','ΦΥΛΟ','ΖΩΗΡΟΣ','ΙΔΙΑΙΤΕΡΟΤΗΤΑ','ΠΑΙΔΙ_ΕΚΠΑΙΔΕΥΤΙΚΟΥ','ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ','ΦΙΛΟΙ']
            step_cols = [f'ΒΗΜΑ1_ΣΕΝΑΡΙΟ_{sid}', f'ΒΗΜΑ2_ΣΕΝΑΡΙΟ_{sid}', f'ΒΗΜΑ3_ΣΕΝΑΡΙΟ_{sid}', f'ΒΗΜΑ4_ΣΕΝΑΡΙΟ_{sid}']
            out_cols = [c for c in base + step_cols if c in updated_df.columns]
            out_df = updated_df[out_cols].copy()
            out_df[col5] = updated_df[col4]
            out_df.to_excel(writer, index=False, sheet_name=sh)
    return out_xlsx_path
