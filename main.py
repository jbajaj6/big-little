import pandas as pd
from collections import defaultdict


FRESHMAN_CAP = 1
SOPHOMORE_CAP = 2


def clean_name(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    return x if x else None


def first_name(full_name):
    name = clean_name(full_name)
    if not name:
        return None
    return name.split()[0]


def load_preferences(
    freshmen_file,
    sophomores_file,
    freshmen_sheet=0,
    sophomores_sheet=0,
):
    freshmen_df = pd.read_excel(freshmen_file, sheet_name=freshmen_sheet)
    sophomores_df = pd.read_excel(sophomores_file, sheet_name=sophomores_sheet)

    freshmen_df.columns = [str(c).strip() for c in freshmen_df.columns]
    sophomores_df.columns = [str(c).strip() for c in sophomores_df.columns]

    # Assumption:
    # col 0 = timestamp
    # col 1 = email
    # col 2 = person's own name
    # col 3 onward = ranked preferences

    freshman_name_col = freshmen_df.columns[2]
    sophomore_name_col = sophomores_df.columns[2]

    freshman_pref_cols = list(freshmen_df.columns[3:])
    sophomore_pref_cols = list(sophomores_df.columns[3:])

    freshman_prefs = {}
    for _, row in freshmen_df.iterrows():
        freshman_name = clean_name(row[freshman_name_col])
        if not freshman_name:
            continue

        prefs = [clean_name(row[col]) for col in freshman_pref_cols]
        prefs = [p for p in prefs if p is not None]
        freshman_prefs[freshman_name] = prefs

    sophomore_prefs = {}
    for _, row in sophomores_df.iterrows():
        sophomore_name = clean_name(row[sophomore_name_col])
        if not sophomore_name:
            continue

        prefs = [clean_name(row[col]) for col in sophomore_pref_cols]
        prefs = [p for p in prefs if p is not None]
        sophomore_prefs[sophomore_name] = prefs

    return freshman_prefs, sophomore_prefs


def build_rank_maps(freshman_prefs, sophomore_prefs):
    freshman_rank = {}
    for f, prefs in freshman_prefs.items():
        freshman_rank[f] = {s: i for i, s in enumerate(prefs)}

    sophomore_rank = {}
    for s, prefs in sophomore_prefs.items():
        sophomore_rank[s] = {f_first: i for i, f_first in enumerate(prefs)}

    return freshman_rank, sophomore_rank


def remove_unavailable_from_lists(
    freshman_prefs_current,
    sophomore_prefs_current,
    freshman_match_count,
    sophomore_match_count,
):
    matched_freshmen = {
        f for f, count in freshman_match_count.items()
        if count >= FRESHMAN_CAP
    }
    full_sophomores = {
        s for s, count in sophomore_match_count.items()
        if count >= SOPHOMORE_CAP
    }

    matched_freshman_first_names = {first_name(f) for f in matched_freshmen}

    for f in freshman_prefs_current:
        freshman_prefs_current[f] = [
            s for s in freshman_prefs_current[f]
            if s not in full_sophomores
        ]

    for s in sophomore_prefs_current:
        sophomore_prefs_current[s] = [
            f_first for f_first in sophomore_prefs_current[s]
            if f_first not in matched_freshman_first_names
        ]


def get_top_available_choice(pref_list):
    return pref_list[0] if pref_list else None


def mutual_top_matches(
    freshman_prefs_current,
    sophomore_prefs_current,
    freshman_match_count,
    sophomore_match_count,
):
    freshman_top = {}
    for f, prefs in freshman_prefs_current.items():
        if freshman_match_count[f] < FRESHMAN_CAP and prefs:
            freshman_top[f] = get_top_available_choice(prefs)

    sophomore_top = {}
    for s, prefs in sophomore_prefs_current.items():
        if sophomore_match_count[s] < SOPHOMORE_CAP and prefs:
            sophomore_top[s] = get_top_available_choice(prefs)

    matches = []
    for f, s in freshman_top.items():
        f_first = first_name(f)
        if s in sophomore_top and sophomore_top[s] == f_first:
            matches.append((f, s))

    return matches


def best_fallback_pair(
    freshman_prefs_current,
    sophomore_prefs_current,
    freshman_match_count,
    sophomore_match_count,
    freshman_rank,
    sophomore_rank,
):
    best_pair = None
    best_score = float("inf")

    for f, prefs in freshman_prefs_current.items():
        if freshman_match_count[f] >= FRESHMAN_CAP:
            continue

        f_first = first_name(f)

        for s in prefs:
            if sophomore_match_count[s] >= SOPHOMORE_CAP:
                continue

            if f_first not in sophomore_rank[s]:
                continue

            score = freshman_rank[f][s] + sophomore_rank[s][f_first]

            if score < best_score:
                best_score = score
                best_pair = (f, s)

    return best_pair


def run_matching(freshman_prefs, sophomore_prefs):
    freshman_prefs_current = {
        f: prefs[:] for f, prefs in freshman_prefs.items()
    }
    sophomore_prefs_current = {
        s: prefs[:] for s, prefs in sophomore_prefs.items()
    }

    freshman_rank, sophomore_rank = build_rank_maps(
        freshman_prefs,
        sophomore_prefs,
    )

    freshman_match_count = defaultdict(int)
    sophomore_match_count = defaultdict(int)
    matches = []

    while True:
        remove_unavailable_from_lists(
            freshman_prefs_current,
            sophomore_prefs_current,
            freshman_match_count,
            sophomore_match_count,
        )

        unmatched_freshmen = [
            f for f in freshman_prefs_current
            if freshman_match_count[f] < FRESHMAN_CAP
        ]

        if not unmatched_freshmen:
            break

        round_matches = mutual_top_matches(
            freshman_prefs_current,
            sophomore_prefs_current,
            freshman_match_count,
            sophomore_match_count,
        )

        used_freshmen = set()
        sophomore_slots_left = {
            s: SOPHOMORE_CAP - sophomore_match_count[s]
            for s in sophomore_prefs_current
        }
        filtered_round_matches = []

        for f, s in round_matches:
            if f in used_freshmen:
                continue
            if sophomore_slots_left.get(s, 0) <= 0:
                continue

            filtered_round_matches.append((f, s))
            used_freshmen.add(f)
            sophomore_slots_left[s] -= 1

        if filtered_round_matches:
            for f, s in filtered_round_matches:
                matches.append((f, s))
                freshman_match_count[f] += 1
                sophomore_match_count[s] += 1
            continue

        pair = best_fallback_pair(
            freshman_prefs_current,
            sophomore_prefs_current,
            freshman_match_count,
            sophomore_match_count,
            freshman_rank,
            sophomore_rank,
        )

        if pair is None:
            break

        f, s = pair
        matches.append((f, s))
        freshman_match_count[f] += 1
        sophomore_match_count[s] += 1

    return matches, freshman_match_count, sophomore_match_count


def save_results(matches, sophomores, output_file="big_little_matches.xlsx"):
    sophomore_to_littles = {s: [] for s in sophomores}

    for freshman, sophomore in matches:
        sophomore_to_littles[sophomore].append(freshman)

    rows = []
    for sophomore in sophomores:
        littles = sophomore_to_littles.get(sophomore, [])
        first_little = littles[0] if len(littles) > 0 else ""
        second_little = littles[1] if len(littles) > 1 else ""

        rows.append({
            "Sophomore": sophomore,
            "First Little": first_little,
            "Second Little": second_little,
        })

    results_df = pd.DataFrame(rows)
    results_df.to_excel(output_file, index=False)
    return results_df


def print_summary(results_df):
    print("\nFinal grouped matches:")
    print(results_df.to_string(index=False))


if __name__ == "__main__":
    freshmen_file = "freshmen_rankings.xlsx"
    sophomores_file = "sophomore_rankings.xlsx"

    freshman_prefs, sophomore_prefs = load_preferences(
        freshmen_file=freshmen_file,
        sophomores_file=sophomores_file,
    )

    matches, freshman_match_count, sophomore_match_count = run_matching(
        freshman_prefs,
        sophomore_prefs,
    )

    results_df = save_results(
        matches=matches,
        sophomores=list(sophomore_prefs.keys()),
        output_file="big_little_matches.xlsx",
    )

    print_summary(results_df)