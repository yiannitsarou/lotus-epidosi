# -*- coding: utf-8 -*-
"""
STEP7.xlsx — Multi-class με λογική "αντίθετων προσήμων" για ΕΠΙΔΟΣΗ 3 επιπέδων
Νέα χαρακτηριστικά:
- Ισορροπία #επιδ_1 & #επιδ_3 με αντίθετα πρόσημα → αυτόματη ισορροπία #επιδ_2
- Veto: spread επίδ_1 ≤2, spread επίδ_3 ≤2, αποφυγή ίδιων προσήμων
- Greedy scoring με bonus για opposite signs
- Κρατάει όλα τα constraints (φύλο, γλώσσα, dyads, flags)
"""
import pandas as pd
import re
from typing import Optional, Tuple, List, Dict, Iterable
import itertools

# ---------------- CONFIG ----------------
STEP3_XLSX = "STEP3_SCENARIOS.xlsx"
STEP6_XLSX = "STEP6_PER_SCENARIO_OUTPUT_FIXED.xlsx"
ROSTER_XLSX = "Παραδειγμα1.xlsx"

SCENARIOS = ["ΣΕΝΑΡΙΟ_1","ΣΕΝΑΡΙΟ_2","ΣΕΝΑΡΙΟ_3","ΣΕΝΑΡΙΟ_4","ΣΕΝΑΡΙΟ_5"]

TARGET_DIFF = 2
MAX_SWAPS = 40

ENABLE_DYAD_SWAPS   = True
ENABLE_SINGLE_SWAPS = True
REQUIRE_SAME_CATEGORY = True
MAX_GENDER_DIFF = 2
MAX_GREEK_DIFF  = 3

OUT_XLSX = "STEP7.xlsx"
# ----------------------------------------

def sgn(x: int) -> int:
    """Πρόσημο: -1, 0, +1"""
    return 0 if x == 0 else (1 if x > 0 else -1)

def _norm_class(val):
    if pd.isna(val): return None
    s = str(val).strip().upper().replace("Α","A")
    s = re.sub(r"[\s\\-]", "", s)
    s = s.replace("01","1").replace("02","2").replace("03","3").replace("04","4").replace("05","5")
    if re.fullmatch(r"A[1-5]", s):
        return s
    return None

def _fmt_like_input(tag: str, greek_style: bool) -> str:
    if not tag: return tag
    if greek_style:
        return tag.replace("A","Α",1)
    return tag

def _find_col(df, options: Tuple[str, ...]):
    for c in df.columns:
        for o in options:
            if str(c).strip().upper() == o.upper():
                return c
    for c in df.columns:
        up = str(c).strip().upper()
        for o in options:
            if o.upper() in up:
                return c
    return None

def _to_int(x):
    try:
        return int(round(float(str(x).replace(",", "."))))
    except:
        return None

def _uses_greek_alpha(series):
    for v in series.dropna().astype(str).head(50):
        if "Α" in v:
            return True
    return False

def _parse_friends_list(s: str):
    if pd.isna(s): return []
    parts = re.split(r"[,;\\n·]+", str(s))
    return [p.strip() for p in parts if p.strip()]

def _norm_gender(x):
    if pd.isna(x): return None
    s = str(x).strip().upper()
    if s.startswith("ΚΟΡΙ"): return "Κορίτσια"
    if s.startswith("ΑΓΟ"): return "Αγόρια"
    if s in {"Κ","ΚΟ","ΚΟΡ"}: return "Κορίτσια"
    if s in {"Α","ΑΓ"}: return "Αγόρια"
    if "GIRL" in s: return "Κορίτσια"
    if "BOY"  in s: return "Αγόρια"
    return s.title()

def _norm_lang(x):
    if pd.isna(x): return None
    s = str(x).strip().upper()
    if s in {"Ν","N","YES","Y"} or "ΝΑΙ" in s: return "Ν"
    if s in {"Ο","O","NO"} or "ΟΧΙ" in s or "OXI" in s: return "Ο"
    return s

def _norm_yes(x):
    if pd.isna(x): return False
    s = str(x).strip().upper()
    return s in {"Ν","N","YES","Y","TRUE","1","ΝΑΙ"}

def _norm_zoiros(x):
    if pd.isna(x): return "Ο"
    s = str(x).strip().upper()
    if s in {"N1","N2","N3"}: s = "Ν" + s[1]
    if s in {"Ν1","Ν2","Ν3"}: return s
    if s in {"Ν","N"}: return "Ν3"
    return s

def _pair_category(genders, langs):
    gset = set(genders)
    lset = set(langs)
    if gset == {"Κορίτσια"}:
        if lset == {"Ν"}: return "Καλή Γνώση (Κορίτσια)"
        if lset == {"Ο"}: return "Όχι Καλή Γνώση (Κορίτσια)"
        return "Μικτής Γνώσης (Κορίτσια)"
    if gset == {"Αγόρια"}:
        if lset == {"Ν"}: return "Καλή Γνώση (Αγόρια)"
        if lset == {"Ο"}: return "Όχι Καλή Γνώση (Αγόρια)"
        return "Μικτής Γνώσης (Αγόρια)"
    return "Ομάδες Μικτού Φύλου"

def build_allowed_dyads_from_step3(step3_xlsx: str, step3_sheet: Optional[str]=None) -> pd.DataFrame:
    x = pd.ExcelFile(step3_xlsx)
    if step3_sheet is None:
        prefers = [s for s in x.sheet_names if re.search(r"ΒΗΜΑ\\s*3|SCENARIO|ΣΕΝΑΡΙΟ", s, re.IGNORECASE)]
        step3_sheet = prefers[0] if prefers else x.sheet_names[0]
    df = x.parse(step3_sheet)

    name_c   = _find_col(df, ("ΟΝΟΜΑ","ΟΝΟΜΑΤΕΠΩΝΥΜΟ","NAME","ΜΑΘΗΤΗΣ"))
    gender_c = _find_col(df, ("ΦΥΛΟ","GENDER"))
    lang_c   = _find_col(df, ("ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ","ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ","LANG"))
    friends_c= _find_col(df, ("ΦΙΛΟΙ","FRIENDS"))
    kid_c    = _find_col(df, ("ΠΑΙΔΙ_ΕΚΠΑΙΔΕΥΤΙΚΟΥ",))
    zoiros_c = _find_col(df, ("ΖΩΗΡΟΣ",))
    spec_c   = _find_col(df, ("ΙΔΙΑΙΤΕΡΟΤΗΤΑ","SPECIAL"))

    if not all([name_c, gender_c, lang_c, friends_c]):
        raise ValueError("Βήμα 3: Λείπουν στήλες ΟΝΟΜΑ/ΦΥΛΟ/ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ/ΦΙΛΟΙ.")

    roster = df[[name_c, gender_c, lang_c, friends_c]].copy()
    roster["__name"]    = roster[name_c].astype(str).str.strip()
    roster["__gender"]  = roster[gender_c].map(_norm_gender)
    roster["__lang"]    = roster[lang_c].map(_norm_lang)
    roster["__friends"] = roster[friends_c].map(_parse_friends_list)

    roster["__kid"]  = df[kid_c].map(_norm_yes) if (kid_c and kid_c in df.columns) else False
    roster["__spec"] = df[spec_c].map(_norm_yes) if (spec_c and spec_c in df.columns) else False
    roster["__zoi"]  = df[zoiros_c].map(_norm_zoiros) if (zoiros_c and zoiros_c in df.columns) else "Ο"

    name_to_idx = {n:i for i,n in enumerate(roster["__name"].tolist())}

    seen = set()
    pairs = []
    for i, row in roster.iterrows():
        me = row["__name"]
        for f in row["__friends"]:
            if f in name_to_idx:
                j = name_to_idx[f]
                if me in roster.loc[j,"__friends"]:
                    a,b = sorted([me, f])
                    key = (a,b)
                    if key in seen: 
                        continue
                    seen.add(key)
                    flags_i = (roster.loc[i,"__kid"] or roster.loc[i,"__spec"] or roster.loc[i,"__zoi"] in {"Ν1","Ν2","Ν3"})
                    flags_j = (roster.loc[j,"__kid"] or roster.loc[j,"__spec"] or roster.loc[j,"__zoi"] in {"Ν1","Ν2","Ν3"})
                    if flags_i or flags_j:
                        continue
                    genders = [roster.loc[i,"__gender"], roster.loc[j,"__gender"]]
                    langs   = [roster.loc[i,"__lang"],   roster.loc[j,"__lang"]]
                    cat = _pair_category(genders, langs)
                    pairs.append((a,b,cat))

    return pd.DataFrame(pairs, columns=["Όνομα 1","Όνομα 2","Κατηγορία"])

def _counts_by_level(classes: pd.Series, perf: pd.Series, level: int) -> Dict[str, int]:
    """Μέτρηση μαθητών με ΕΠΙΔΟΣΗ=level ανά τμήμα"""
    counts = {}
    for c in sorted(set(classes.dropna())):
        counts[c] = int(((classes==c) & (perf==level)).sum())
    return counts

def _class_style_tag(c: str, greek: bool) -> str:
    return _fmt_like_input(c, greek)

def _category_key(gender:str, greek:str) -> Tuple[str,str]:
    g = (gender or "").strip().title()
    if g.upper().startswith("ΑΓ"): g = "Αγόρια"
    elif g.upper().startswith("ΚΟ"): g = "Κορίτσια"
    l = "Ν" if (str(greek or "").strip().upper()=="Ν") else "Ο"
    return (g, l)

def _same_category(g1:str, l1:str, g2:str, l2:str) -> bool:
    if not REQUIRE_SAME_CATEGORY:
        return True
    return _category_key(g1,l1) == _category_key(g2,l2)

def _counts_by_pred(df: pd.DataFrame, class_col: str, pred) -> pd.Series:
    tmp = df[df.apply(pred, axis=1)]
    res = tmp.groupby(class_col)["AM"].count() if "AM" in tmp.columns else tmp.groupby(class_col).size()
    all_classes = df[class_col].dropna().unique().tolist()
    return res.reindex(sorted(all_classes)).fillna(0).astype(int)

def _violates_balance_after_swap(df: pd.DataFrame, class_col: str,
                                 moves_out: Iterable[str], cls_out: str,
                                 moves_in: Iterable[str], cls_in: str) -> bool:
    """Veto: φύλο/γλώσσα spread + αντίθετα πρόσημα για ΕΠΙΔΟΣΗ"""
    if not MAX_GENDER_DIFF and not MAX_GREEK_DIFF:
        return False
    
    tmp = df[["AM","ΦΥΛΟ","ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ","ΕΠΙΔΟΣΗ", class_col]].copy()
    if moves_out:
        tmp.loc[tmp["AM"].isin(list(moves_out)) & (tmp[class_col]==cls_out), class_col] = "__TMP__"
    if moves_in:
        tmp.loc[tmp["AM"].isin(list(moves_in)) & (tmp[class_col]==cls_in), class_col] = "__IN__"
    tmp.loc[tmp[class_col]=="__TMP__", class_col] = cls_in
    tmp.loc[tmp[class_col]=="__IN__",  class_col] = cls_out

    # Φύλο/Γλώσσα veto (υπάρχον)
    boys  = _counts_by_pred(tmp, class_col, lambda r: str(r["ΦΥΛΟ"]).strip().upper().startswith("ΑΓ"))
    girls = _counts_by_pred(tmp, class_col, lambda r: str(r["ΦΥΛΟ"]).strip().upper().startswith("ΚΟ"))
    ggood = _counts_by_pred(tmp, class_col, lambda r: str(r["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"]).strip().upper()=="Ν")
    gpoor = _counts_by_pred(tmp, class_col, lambda r: str(r["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"]).strip().upper()!="Ν")

    if MAX_GENDER_DIFF is not None:
        if (boys.max()-boys.min()) > MAX_GENDER_DIFF:   return True
        if (girls.max()-girls.min()) > MAX_GENDER_DIFF: return True
    if MAX_GREEK_DIFF is not None:
        if (ggood.max()-ggood.min()) > MAX_GREEK_DIFF:  return True
        if (gpoor.max()-gpoor.min()) > MAX_GREEK_DIFF:  return True

    # ΕΠΙΔΟΣΗ: Veto για αντίθετα πρόσημα
    epid_1 = _counts_by_pred(tmp, class_col, lambda r: r["ΕΠΙΔΟΣΗ"]==1)
    epid_3 = _counts_by_pred(tmp, class_col, lambda r: r["ΕΠΙΔΟΣΗ"]==3)
    
    # Hard veto: spread ≤2
    if (epid_1.max() - epid_1.min()) > 2: return True
    if (epid_3.max() - epid_3.min()) > 2: return True
    
    # Soft veto: Αν |Δ1|>1 ΚΑΙ |Δ3|>1 με ίδια πρόσημα → veto
    all_classes = sorted(tmp[class_col].dropna().unique())
    for cls_a, cls_b in itertools.combinations(all_classes, 2):
        d1 = int(epid_1[cls_a] - epid_1[cls_b])
        d3 = int(epid_3[cls_a] - epid_3[cls_b])
        if abs(d1) > 1 and abs(d3) > 1:
            if sgn(d1) == sgn(d3):  # ίδια πρόσημα
                return True
    
    return False

def run_step7_for_scenario(step6_xlsx: str, roster_xlsx: str, step3_xlsx: str,
                           scenario_name: str, step3_sheet: Optional[str]=None,
                           target_diff: int=2, max_swaps: int=40):
    s6 = pd.ExcelFile(step6_xlsx)
    if scenario_name not in s6.sheet_names:
        raise ValueError(f"Το φύλλο {scenario_name} δεν υπάρχει στο {step6_xlsx}")
    df = s6.parse(scenario_name)

    name_col  = _find_col(df, ("ΟΝΟΜΑ","ΟΝΟΜΑΤΕΠΩΝΥΜΟ","NAME","ΜΑΘΗΤΗΣ"))
    class_col = _find_col(df, (f"ΒΗΜΑ6_{scenario_name}",)) or _find_col(df, ("ΤΜΗΜΑ","CLASS","SECTION"))
    am_col    = _find_col(df, ("AM","ΑΜ","ID"))
    if not name_col or not class_col:
        raise ValueError(f"{scenario_name}: δεν εντοπίστηκαν στήλες Ονόματος/Τμήματος.")
    if not am_col:
        df["AM"] = df[name_col].astype(str).str.strip()
        am_col = "AM"

    classes0 = df[class_col].apply(_norm_class)
    names    = df[name_col].astype(str).str.strip()
    AMs      = df[am_col].astype(str).str.strip()
    greek_style = _uses_greek_alpha(df[class_col].astype(str))

    roster = pd.ExcelFile(roster_xlsx).parse(0)
    r_name = _find_col(roster, ("ΟΝΟΜΑ","ΟΝΟΜΑΤΕΠΩΝΥΜΟ","NAME"))
    r_perf = _find_col(roster, ("ΕΠΙΔΟΣΗ",))
    r_gen  = _find_col(roster, ("ΦΥΛΟ","GENDER"))
    r_lang = _find_col(roster, ("ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ","ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ","LANG"))

    if not r_name or not r_perf:
        raise ValueError("Roster: λείπουν στήλες ΟΝΟΜΑ/ΕΠΙΔΟΣΗ.")
    perf_map  = {str(n).strip(): _to_int(p) for n,p in roster[[r_name, r_perf]].dropna().values}
    gender_map= {str(n).strip(): _norm_gender(g) for n,g in roster[[r_name, r_gen]].dropna().values} if r_gen else {}
    greek_map = {str(n).strip(): _norm_lang(l)   for n,l in roster[[r_name, r_lang]].dropna().values} if r_lang else {}

    perf = names.map(lambda n: perf_map.get(n))

    valid = classes0.notna()
    classes = classes0.where(valid)
    names   = names.where(valid)
    perf    = perf.where(valid)
    AMs     = AMs.where(valid)

    # Baseline counts για επίπεδα 1 & 3
    counts_1 = _counts_by_level(classes, pd.Series(perf), 1)
    counts_3 = _counts_by_level(classes, pd.Series(perf), 3)
    if not counts_1:
        raise ValueError(f"{scenario_name}: δεν βρέθηκαν έγκυρα τμήματα.")
    spread_1_before = max(counts_1.values()) - min(counts_1.values())
    spread_3_before = max(counts_3.values()) - min(counts_3.values())

    dy = build_allowed_dyads_from_step3(step3_xlsx, step3_sheet)

    def _pair_perf_counts(a, b):
        pa = perf_map.get(str(a).strip())
        pb = perf_map.get(str(b).strip())
        return {
            "c1": int((pa==1) + (pb==1)),
            "c3": int((pa==3) + (pb==3))
        }
    
    dy["perf_counts"] = [_pair_perf_counts(a,b) for a,b in zip(dy["Όνομα 1"], dy["Όνομα 2"])]

    name_to_class = dict(zip(names.fillna("").tolist(), classes.fillna("").tolist()))

    def _dy_class(row):
        a = str(row["Όνομα 1"]).strip(); b = str(row["Όνομα 2"]).strip()
        ca = name_to_class.get(a); cb = name_to_class.get(b)
        return ca if (ca in {"A1","A2","A3","A4","A5"} and ca == cb) else None

    dy["class"] = dy.apply(_dy_class, axis=1)
    dy_in = dy.dropna(subset=["class"]).copy()

    work = pd.DataFrame({
        "AM": AMs,
        "ΟΝΟΜΑ": names,
        "ΦΥΛΟ": names.map(lambda n: gender_map.get(n, "")),
        "ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ": names.map(lambda n: greek_map.get(n, "")),
        "ΕΠΙΔΟΣΗ": perf,
        "CLASS": classes
    })

    b7 = classes.copy().map(lambda c: _class_style_tag(c, greek_style))

    swaps = []
    cur_counts_1 = counts_1.copy()
    cur_counts_3 = counts_3.copy()

    def pick_classes_to_swap(c1: Dict[str,int], c3: Dict[str,int]):
        # Επιλογή με βάση συνδυασμό spread 1 & 3
        items_1 = sorted(c1.items(), key=lambda kv: kv[1])
        items_3 = sorted(c3.items(), key=lambda kv: kv[1])
        
        # Donor: max στο 1 ή max στο 3
        donor_by_1 = items_1[-1][0]
        donor_by_3 = items_3[-1][0]
        receiver_by_1 = items_1[0][0]
        receiver_by_3 = items_3[0][0]
        
        # Προτιμάμε το ζεύγος με μεγαλύτερο συνολικό spread
        spread_opt1 = (c1[donor_by_1] - c1[receiver_by_1]) + (c3[donor_by_1] - c3[receiver_by_1])
        spread_opt2 = (c1[donor_by_3] - c1[receiver_by_3]) + (c3[donor_by_3] - c3[receiver_by_3])
        
        if spread_opt1 >= spread_opt2:
            return donor_by_1, receiver_by_1
        return donor_by_3, receiver_by_3

    swaps_done = 0

    # -------------------- ΦΑΣΗ Α: ΔΥΑΔΕΣ --------------------
    if ENABLE_DYAD_SWAPS:
        for _ in range(max_swaps):
            spread_1 = max(cur_counts_1.values()) - min(cur_counts_1.values())
            spread_3 = max(cur_counts_3.values()) - min(cur_counts_3.values())
            if spread_1 <= target_diff and spread_3 <= target_diff:
                break
            if swaps_done >= max_swaps:
                break
            
            donor, receiver = pick_classes_to_swap(cur_counts_1, cur_counts_3)

            candidates = []
            cats = sorted(set(dy_in.loc[dy_in["class"]==donor,"Κατηγορία"]) & 
                         set(dy_in.loc[dy_in["class"]==receiver,"Κατηγορία"]))
            if not cats:
                break

            for cat in cats:
                L = dy_in[(dy_in["class"]==donor) & (dy_in["Κατηγορία"]==cat)]
                R = dy_in[(dy_in["class"]==receiver) & (dy_in["Κατηγορία"]==cat)]
                for _, li in L.iterrows():
                    for _, rj in R.iterrows():
                        donor_names    = [str(li["Όνομα 1"]).strip(), str(li["Όνομα 2"]).strip()]
                        receiver_names = [str(rj["Όνομα 1"]).strip(), str(rj["Όνομα 2"]).strip()]

                        # Υπολογισμός post-swap counts
                        li_c = li["perf_counts"]
                        rj_c = rj["perf_counts"]
                        
                        post_donor_1 = cur_counts_1[donor] - li_c["c1"] + rj_c["c1"]
                        post_recv_1  = cur_counts_1[receiver] - rj_c["c1"] + li_c["c1"]
                        post_donor_3 = cur_counts_3[donor] - li_c["c3"] + rj_c["c3"]
                        post_recv_3  = cur_counts_3[receiver] - rj_c["c3"] + li_c["c3"]
                        
                        trial_1 = cur_counts_1.copy()
                        trial_1[donor] = int(post_donor_1)
                        trial_1[receiver] = int(post_recv_1)
                        
                        trial_3 = cur_counts_3.copy()
                        trial_3[donor] = int(post_donor_3)
                        trial_3[receiver] = int(post_recv_3)
                        
                        new_spread_1 = max(trial_1.values()) - min(trial_1.values())
                        new_spread_3 = max(trial_3.values()) - min(trial_3.values())

                        # Veto
                        if _violates_balance_after_swap(work.assign(CLASS=b7.map(lambda x: x.replace("Α","A",1))),
                                                        "CLASS", donor_names, donor, receiver_names, receiver):
                            continue

                        # Scoring με bonus για αντίθετα πρόσημα
                        diff_1 = int(post_donor_1 - post_recv_1)
                        diff_3 = int(post_donor_3 - post_recv_3)
                        opposite_bonus = 50 if (sgn(diff_1) * sgn(diff_3) < 0) else 0
                        
                        score = -100 * (new_spread_1 + new_spread_3) + opposite_bonus

                        candidates.append({
                            "Κατηγορία": cat,
                            "donor_pair": donor_names,
                            "receiver_pair": receiver_names,
                            "donor_c1": li_c["c1"], "donor_c3": li_c["c3"],
                            "receiver_c1": rj_c["c1"], "receiver_c3": rj_c["c3"],
                            "new_spread_1": int(new_spread_1),
                            "new_spread_3": int(new_spread_3),
                            "post_donor_1": int(post_donor_1), "post_recv_1": int(post_recv_1),
                            "post_donor_3": int(post_donor_3), "post_recv_3": int(post_recv_3),
                            "opposite_signs": opposite_bonus > 0,
                            "score": score
                        })
            
            if not candidates:
                break
            
            cdf = pd.DataFrame(candidates).sort_values(by=["score"], ascending=False)
            best = cdf.iloc[0].to_dict()

            donor_names    = list(best["donor_pair"])
            receiver_names = list(best["receiver_pair"])

            mask_d = names.isin(donor_names)
            mask_r = names.isin(receiver_names)
            b7.loc[mask_d] = _class_style_tag(receiver, greek_style)
            b7.loc[mask_r] = _class_style_tag(donor, greek_style)

            swaps.append({
                "Σενάριο": scenario_name,
                "Τύπος": "Δυάδες",
                "Από": donor, "Προς": receiver,
                "Κατηγορία": best["Κατηγορία"],
                "Ζεύγος από Δότη": f"{donor_names[0]} — {donor_names[1]} (επιδ: 1={best['donor_c1']}, 3={best['donor_c3']})",
                "Ζεύγος από Δέκτη": f"{receiver_names[0]} — {receiver_names[1]} (επιδ: 1={best['receiver_c1']}, 3={best['receiver_c3']})",
                "Spread_1 μετά": best["new_spread_1"],
                "Spread_3 μετά": best["new_spread_3"],
                "Αντίθετα Πρόσημα": "✓" if best["opposite_signs"] else "✗"
            })

            cur_counts_1[donor] = best["post_donor_1"]
            cur_counts_1[receiver] = best["post_recv_1"]
            cur_counts_3[donor] = best["post_donor_3"]
            cur_counts_3[receiver] = best["post_recv_3"]
            swaps_done += 1

            name_to_class.update({donor_names[0]: receiver, donor_names[1]: receiver,
                                  receiver_names[0]: donor,  receiver_names[1]: donor})
            def _upd(row):
                a = str(row["Όνομα 1"]).strip(); b = str(row["Όνομα 2"]).strip()
                ca = name_to_class.get(a); cb = name_to_class.get(b)
                return ca if (ca in {"A1","A2","A3","A4","A5"} and ca == cb) else None
            dy_in["class"] = dy_in.apply(_upd, axis=1)

    # -------------------- ΦΑΣΗ Β: ΜΕΜΟΝΩΜΕΝΟΙ --------------------
    if ENABLE_SINGLE_SWAPS and swaps_done < max_swaps:
        dy_names = set(dy["Όνομα 1"]).union(set(dy["Όνομα 2"]))
        work2 = pd.DataFrame({
            "AM": AMs,
            "ΟΝΟΜΑ": names,
            "ΦΥΛΟ": work["ΦΥΛΟ"],
            "ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ": work["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"],
            "ΕΠΙΔΟΣΗ": perf,
            "CLASS": b7.map(lambda x: x.replace("Α","A",1))
        })
        work2["eligible_single"] = ~work2["ΟΝΟΜΑ"].isin(dy_names)

        def pick_single_candidates(donor:str, receiver:str):
            # Από donor: singles με ΕΠΙΔΟΣΗ=1 ή 3
            Sd_1 = work2[(work2["CLASS"]==donor) & (work2["eligible_single"]) & (work2["ΕΠΙΔΟΣΗ"]==1)].copy()
            Sd_3 = work2[(work2["CLASS"]==donor) & (work2["eligible_single"]) & (work2["ΕΠΙΔΟΣΗ"]==3)].copy()
            # Από receiver: singles με ΕΠΙΔΟΣΗ=1 ή 3
            Sr_1 = work2[(work2["CLASS"]==receiver) & (work2["eligible_single"]) & (work2["ΕΠΙΔΟΣΗ"]==1)].copy()
            Sr_3 = work2[(work2["CLASS"]==receiver) & (work2["eligible_single"]) & (work2["ΕΠΙΔΟΣΗ"]==3)].copy()
            return Sd_1, Sd_3, Sr_1, Sr_3

        while swaps_done < max_swaps:
            spread_1 = max(cur_counts_1.values()) - min(cur_counts_1.values())
            spread_3 = max(cur_counts_3.values()) - min(cur_counts_3.values())
            if spread_1 <= target_diff and spread_3 <= target_diff:
                break
            
            donor, receiver = pick_classes_to_swap(cur_counts_1, cur_counts_3)

            Sd_1, Sd_3, Sr_1, Sr_3 = pick_single_candidates(donor, receiver)

            cands = []
            
            # Swap type 1: 1↔3 (donor has excess 1, receiver has excess 3)
            if not Sd_1.empty and not Sr_3.empty:
                for _, a in Sd_1.iterrows():
                    for _, b in Sr_3.iterrows():
                        if not _same_category(a["ΦΥΛΟ"], a["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"], 
                                             b["ΦΥΛΟ"], b["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"]):
                            continue

                        post_d1 = cur_counts_1[donor] - 1
                        post_r1 = cur_counts_1[receiver] + 1
                        post_d3 = cur_counts_3[donor] + 1
                        post_r3 = cur_counts_3[receiver] - 1
                        
                        trial_1 = cur_counts_1.copy()
                        trial_1[donor] = post_d1
                        trial_1[receiver] = post_r1
                        
                        trial_3 = cur_counts_3.copy()
                        trial_3[donor] = post_d3
                        trial_3[receiver] = post_r3
                        
                        new_sp1 = max(trial_1.values()) - min(trial_1.values())
                        new_sp3 = max(trial_3.values()) - min(trial_3.values())

                        if _violates_balance_after_swap(work2, "CLASS", [a["AM"]], donor, [b["AM"]], receiver):
                            continue

                        diff_1 = post_d1 - post_r1
                        diff_3 = post_d3 - post_r3
                        opposite_bonus = 50 if (sgn(diff_1) * sgn(diff_3) < 0) else 0
                        score = -100 * (new_sp1 + new_sp3) + opposite_bonus

                        cands.append({
                            "swap_type": "1↔3",
                            "a_AM": a["AM"], "a_name": a["ΟΝΟΜΑ"], "a_perf": 1,
                            "b_AM": b["AM"], "b_name": b["ΟΝΟΜΑ"], "b_perf": 3,
                            "new_spread_1": int(new_sp1), "new_spread_3": int(new_sp3),
                            "post_donor_1": int(post_d1), "post_recv_1": int(post_r1),
                            "post_donor_3": int(post_d3), "post_recv_3": int(post_r3),
                            "opposite_signs": opposite_bonus > 0,
                            "score": score
                        })
            
            # Swap type 2: 3↔1 (donor has excess 3, receiver has excess 1)
            if not Sd_3.empty and not Sr_1.empty:
                for _, a in Sd_3.iterrows():
                    for _, b in Sr_1.iterrows():
                        if not _same_category(a["ΦΥΛΟ"], a["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"], 
                                             b["ΦΥΛΟ"], b["ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ"]):
                            continue

                        post_d1 = cur_counts_1[donor] + 1
                        post_r1 = cur_counts_1[receiver] - 1
                        post_d3 = cur_counts_3[donor] - 1
                        post_r3 = cur_counts_3[receiver] + 1
                        
                        trial_1 = cur_counts_1.copy()
                        trial_1[donor] = post_d1
                        trial_1[receiver] = post_r1
                        
                        trial_3 = cur_counts_3.copy()
                        trial_3[donor] = post_d3
                        trial_3[receiver] = post_r3
                        
                        new_sp1 = max(trial_1.values()) - min(trial_1.values())
                        new_sp3 = max(trial_3.values()) - min(trial_3.values())

                        if _violates_balance_after_swap(work2, "CLASS", [a["AM"]], donor, [b["AM"]], receiver):
                            continue

                        diff_1 = post_d1 - post_r1
                        diff_3 = post_d3 - post_r3
                        opposite_bonus = 50 if (sgn(diff_1) * sgn(diff_3) < 0) else 0
                        score = -100 * (new_sp1 + new_sp3) + opposite_bonus

                        cands.append({
                            "swap_type": "3↔1",
                            "a_AM": a["AM"], "a_name": a["ΟΝΟΜΑ"], "a_perf": 3,
                            "b_AM": b["AM"], "b_name": b["ΟΝΟΜΑ"], "b_perf": 1,
                            "new_spread_1": int(new_sp1), "new_spread_3": int(new_sp3),
                            "post_donor_1": int(post_d1), "post_recv_1": int(post_r1),
                            "post_donor_3": int(post_d3), "post_recv_3": int(post_r3),
                            "opposite_signs": opposite_bonus > 0,
                            "score": score
                        })
            
            if not cands:
                break
            
            cdf = pd.DataFrame(cands).sort_values(by=["score"], ascending=False)
            best = cdf.iloc[0].to_dict()

            b7.loc[names==best["a_name"]] = _class_style_tag(receiver, greek_style)
            b7.loc[names==best["b_name"]] = _class_style_tag(donor, greek_style)

            swaps.append({
                "Σενάριο": scenario_name,
                "Τύπος": "Μεμονωμένοι",
                "Από": donor, "Προς": receiver,
                "Κατηγορία": f"{_category_key(work2.loc[work2['ΟΝΟΜΑ']==best['a_name'],'ΦΥΛΟ'].iloc[0], work2.loc[work2['ΟΝΟΜΑ']==best['a_name'],'ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ'].iloc[0])}",
                "Swap": f"{best['swap_type']}: {best['a_name']}(επιδ={best['a_perf']}) ↔ {best['b_name']}(επιδ={best['b_perf']})",
                "Spread_1 μετά": best["new_spread_1"],
                "Spread_3 μετά": best["new_spread_3"],
                "Αντίθετα Πρόσημα": "✓" if best["opposite_signs"] else "✗"
            })

            cur_counts_1[donor] = best["post_donor_1"]
            cur_counts_1[receiver] = best["post_recv_1"]
            cur_counts_3[donor] = best["post_donor_3"]
            cur_counts_3[receiver] = best["post_recv_3"]
            swaps_done += 1

            work2.loc[work2["ΟΝΟΜΑ"]==best["a_name"], "CLASS"] = receiver
            work2.loc[work2["ΟΝΟΜΑ"]==best["b_name"], "CLASS"] = donor

    # Build output
    out_df = df.copy()
    class_col_used = _find_col(df, (f"ΒΗΜΑ6_{scenario_name}",)) or _find_col(df, ("ΤΜΗΜΑ","CLASS","SECTION"))
    insert_at = list(out_df.columns).index(class_col_used) + 1
    out_df.insert(insert_at, f"ΒΗΜΑ7_{scenario_name}", b7)

    # Summaries με 3 επίπεδα
    counts_2 = _counts_by_level(classes, pd.Series(perf), 2)
    cur_counts_2 = _counts_by_level(b7.map(lambda x: x.replace("Α","A",1) if x else None), 
                                     pd.Series(perf), 2)
    
    summary_rows = []
    for cls in sorted(counts_1.keys()):
        summary_rows.append({
            "Τμήμα": cls,
            "Επίδ_1 (πριν)": counts_1.get(cls, 0),
            "Επίδ_2 (πριν)": counts_2.get(cls, 0),
            "Επίδ_3 (πριν)": counts_3.get(cls, 0),
            "Επίδ_1 (μετά)": cur_counts_1.get(cls, 0),
            "Επίδ_2 (μετά)": cur_counts_2.get(cls, 0),
            "Επίδ_3 (μετά)": cur_counts_3.get(cls, 0)
        })
    
    counts_df = pd.DataFrame(summary_rows)
    
    summary = pd.DataFrame([{
        "Σενάριο": scenario_name,
        "Τμήματα": ", ".join(sorted(counts_1.keys())),
        "Spread_1 (πριν)": spread_1_before,
        "Spread_3 (πριν)": spread_3_before,
        "Spread_1 (μετά)": max(cur_counts_1.values()) - min(cur_counts_1.values()),
        "Spread_3 (μετά)": max(cur_counts_3.values()) - min(cur_counts_3.values()),
        "Spread_2 (μετά)": max(cur_counts_2.values()) - min(cur_counts_2.values()),
        "Στόχος ≤": target_diff,
        "Ανταλλαγές": len(swaps)
    }])
    
    swaps_log = pd.DataFrame(swaps) if swaps else pd.DataFrame()

    return out_df, summary, counts_df, swaps_log

if __name__ == "__main__":
    all_summ = []
    all_counts = []
    all_swaps = []

    s6 = pd.ExcelFile(STEP6_XLSX)
    with pd.ExcelWriter(OUT_XLSX) as writer:
        for scen in SCENARIOS:
            if scen in s6.sheet_names:
                try:
                    out_df, summary, counts_df, swaps_log = run_step7_for_scenario(
                        STEP6_XLSX, ROSTER_XLSX, STEP3_XLSX, scen, None, TARGET_DIFF, MAX_SWAPS
                    )
                    out_df.to_excel(writer, index=False, sheet_name=scen)
                    all_summ.append(summary)
                    counts_df["Σενάριο"] = scen
                    all_counts.append(counts_df)
                    if not swaps_log.empty:
                        all_swaps.append(swaps_log)
                except Exception as e:
                    pd.DataFrame([{"Σενάριο": scen, "Σφάλμα": str(e)}]).to_excel(writer, index=False, sheet_name=scen)
            else:
                pd.DataFrame([{"Σενάριο": scen, "Σφάλμα": "Δεν βρέθηκε στο STEP6"}]).to_excel(writer, index=False, sheet_name=scen)

        if all_summ:
            pd.concat(all_summ, ignore_index=True).to_excel(writer, index=False, sheet_name="Σύνοψη_ΒΗΜΑ7")
        if all_counts:
            pd.concat(all_counts, ignore_index=True).to_excel(writer, index=False, sheet_name="Μετρήσεις_Επίδοσης")
        if all_swaps:
            pd.concat(all_swaps, ignore_index=True).to_excel(writer, index=False, sheet_name="Ανταλλαγές")
                "