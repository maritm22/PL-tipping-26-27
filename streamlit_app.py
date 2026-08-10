from __future__ import annotations

import html
import io
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="PL-Tippen 2026/27", page_icon="🏆", layout="wide")

SCORING = {0: 10, 1: 7, 2: 5, 3: 3}
BONUS_POINTS = 15
IMAGE_DIR = Path("images")
FPL_BOOTSTRAP = "https://fantasy.premierleague.com/api/bootstrap-static/"
FPL_FIXTURES = "https://fantasy.premierleague.com/api/fixtures/"

TEAM_ALIASES = {
    "brighton": "Brighton & Hove Albion",
    "brighton & hove albion": "Brighton & Hove Albion",
    "man city": "Manchester City",
    "manchester city": "Manchester City",
    "man utd": "Manchester United",
    "manchester united": "Manchester United",
    "nott'm forest": "Nottingham Forest",
    "nottingham forest": "Nottingham Forest",
    "spurs": "Tottenham Hotspur",
    "tottenham": "Tottenham Hotspur",
    "tottenham hotspur": "Tottenham Hotspur",
    "newcastle": "Newcastle United",
    "newcastle united": "Newcastle United",
    "leeds": "Leeds United",
    "leeds united": "Leeds United",
}

CATEGORY_KEYS = {
    "lag med flest scorede mal": "goals",
    "lag med flest clean sheets": "clean_sheets",
    "lag med flest gule kort": "yellow_cards",
    "lag med flest røde kort": "red_cards",
    "lag med flest uavgjorte kamper": "draws",
    "toppscorer": "top_scorer",
    "flest assist": "assists",
}

CATEGORY_LABELS = {
    "goals": "Flest mål",
    "clean_sheets": "Clean sheets",
    "yellow_cards": "Gule kort",
    "red_cards": "Røde kort",
    "draws": "Uavgjorte",
    "top_scorer": "Toppscorer",
    "assists": "Flest assist",
}


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def norm(value: Any) -> str:
    text = clean_text(value).casefold()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def canonical_team(name: Any) -> str:
    raw = clean_text(name)
    return TEAM_ALIASES.get(norm(raw), raw)


def slug_name(name: str) -> str:
    text = norm(name)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def filename_person_name(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    stem = re.sub(r"^Premier\s*League[-\s]*tips\s*[–—-]?\s*", "", stem, flags=re.I)
    return stem.strip()


def find_row_with_label(ws, label: str, col: int = 1) -> int | None:
    target = norm(label)
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, col).value) == target:
            return r
    return None


def parse_tip_file(uploaded) -> dict[str, Any]:
    filename = getattr(uploaded, "name", "Ukjent.xlsx")
    raw = uploaded.getvalue() if hasattr(uploaded, "getvalue") else uploaded.read()
    wb = load_workbook(io.BytesIO(raw), data_only=False, read_only=False)

    required = {"Tabelltips", "Bonustips", "Lagliste"}
    missing_sheets = sorted(required - set(wb.sheetnames))
    if missing_sheets:
        raise ValueError(f"Mangler ark: {', '.join(missing_sheets)}")

    tab = wb["Tabelltips"]
    bonus = wb["Bonustips"]
    lag = wb["Lagliste"]

    name_row = find_row_with_label(tab, "Deltakerens navn")
    name = clean_text(tab.cell(name_row, 2).value) if name_row else ""

    team_header = None
    for r in range(1, tab.max_row + 1):
        if norm(tab.cell(r, 1).value) == "tippet plass" and norm(tab.cell(r, 2).value) == "lag":
            team_header = r
            break
    if team_header is None:
        raise ValueError("Fant ikke tabelltips-tabellen.")

    table_tips = []
    for r in range(team_header + 1, tab.max_row + 1):
        place = tab.cell(r, 1).value
        team = clean_text(tab.cell(r, 2).value)
        if isinstance(place, (int, float)) and 1 <= int(place) <= 20:
            table_tips.append({"tip_place": int(place), "team": canonical_team(team)})

    teams = []
    for r in range(2, lag.max_row + 1):
        team = clean_text(lag.cell(r, 1).value)
        if team:
            teams.append(canonical_team(team))

    bonus_header = None
    for r in range(1, bonus.max_row + 1):
        if norm(bonus.cell(r, 2).value) == "kategori" and norm(bonus.cell(r, 3).value) == "tips":
            bonus_header = r
            break
    if bonus_header is None:
        raise ValueError("Fant ikke bonustips-tabellen.")

    bonus_tips = []
    for r in range(bonus_header + 1, bonus.max_row + 1):
        category = clean_text(bonus.cell(r, 2).value)
        pick = clean_text(bonus.cell(r, 3).value)
        answer_type = clean_text(bonus.cell(r, 4).value)
        if category:
            bonus_tips.append({
                "category": category,
                "key": CATEGORY_KEYS.get(norm(category)),
                "pick": canonical_team(pick) if norm(answer_type) == "lag" else pick,
                "answer_type": answer_type,
            })

    warnings = []
    errors = []
    expected_filename_name = filename_person_name(filename)
    if not name:
        errors.append("Deltakernavn mangler.")
    elif expected_filename_name and norm(expected_filename_name) != norm(name):
        warnings.append(f'Filnavnet tilsier «{expected_filename_name}», men arket sier «{name}».')

    tip_teams = [x["team"] for x in table_tips if x["team"]]
    if len(table_tips) != 20:
        errors.append(f"Fant {len(table_tips)} tabelltips, forventet 20.")
    if len(set(map(norm, tip_teams))) != len(tip_teams):
        errors.append("Samme lag er brukt flere ganger i tabelltipsene.")
    missing_team_tips = [t for t in teams if norm(t) not in {norm(x) for x in tip_teams}]
    if missing_team_tips:
        errors.append("Lag som ikke er tippet: " + ", ".join(missing_team_tips))
    blanks = [b["category"] for b in bonus_tips if not b["pick"]]
    if blanks:
        errors.append("Tomme bonustips: " + ", ".join(blanks))
    unknown_categories = [b["category"] for b in bonus_tips if not b["key"]]
    if unknown_categories:
        warnings.append("Ukjente bonuskategorier: " + ", ".join(unknown_categories))

    return {
        "filename": filename,
        "name": name or expected_filename_name or filename,
        "teams": teams,
        "table_tips": table_tips,
        "bonus_tips": bonus_tips,
        "warnings": warnings,
        "errors": errors,
    }


@st.cache_data(ttl=900, show_spinner=False)
def fetch_fpl_data() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    headers = {"User-Agent": "PL-Tippen-Office-League/1.0"}
    b = requests.get(FPL_BOOTSTRAP, headers=headers, timeout=12)
    f = requests.get(FPL_FIXTURES, headers=headers, timeout=12)
    b.raise_for_status()
    f.raise_for_status()
    return b.json(), f.json()


def blank_team_stats(teams: list[str]) -> dict[str, dict[str, int]]:
    return {
        team: {"played": 0, "won": 0, "drawn": 0, "lost": 0, "gf": 0, "ga": 0,
               "gd": 0, "points": 0, "clean_sheets": 0, "yellow_cards": 0, "red_cards": 0}
        for team in teams
    }


def build_live_state(teams: list[str]) -> dict[str, Any]:
    # Always produces a valid GW0 fallback, even if FPL is temporarily unavailable.
    stats = blank_team_stats(teams)
    state = {
        "label": "GW0 – alfabetisk starttabell",
        "source_status": "GW0",
        "stats": stats,
        "table": [],
        "bonus_leaders": {k: {"leaders": set(), "value": 0, "active": False} for k in CATEGORY_LABELS},
        "api_ok": False,
        "api_message": None,
    }

    try:
        bootstrap, fixtures = fetch_fpl_data()
        state["api_ok"] = True
    except Exception as exc:
        bootstrap, fixtures = {}, []
        state["api_message"] = f"FPL-data kunne ikke hentes akkurat nå ({type(exc).__name__}). Viser GW0-fallback."

    fpl_team_id_to_name: dict[int, str] = {}
    if bootstrap:
        for t in bootstrap.get("teams", []):
            fpl_team_id_to_name[int(t["id"])] = canonical_team(t.get("name") or t.get("short_name"))

    relevant_finished = []
    for fx in fixtures:
        if not fx.get("finished"):
            continue
        home = fpl_team_id_to_name.get(int(fx.get("team_h", -1)))
        away = fpl_team_id_to_name.get(int(fx.get("team_a", -1)))
        if home in stats and away in stats:
            relevant_finished.append(fx)
            hg = int(fx.get("team_h_score") or 0)
            ag = int(fx.get("team_a_score") or 0)
            for team, gf, ga in ((home, hg, ag), (away, ag, hg)):
                stats[team]["played"] += 1
                stats[team]["gf"] += gf
                stats[team]["ga"] += ga
                if ga == 0:
                    stats[team]["clean_sheets"] += 1
            if hg > ag:
                stats[home]["won"] += 1; stats[home]["points"] += 3; stats[away]["lost"] += 1
            elif ag > hg:
                stats[away]["won"] += 1; stats[away]["points"] += 3; stats[home]["lost"] += 1
            else:
                stats[home]["drawn"] += 1; stats[away]["drawn"] += 1
                stats[home]["points"] += 1; stats[away]["points"] += 1

    for team in stats:
        stats[team]["gd"] = stats[team]["gf"] - stats[team]["ga"]

    # Card totals from FPL player data.
    player_rows = []
    if bootstrap:
        for e in bootstrap.get("elements", []):
            team_name = fpl_team_id_to_name.get(int(e.get("team", -1)))
            if team_name not in stats:
                continue
            stats[team_name]["yellow_cards"] += int(e.get("yellow_cards") or 0)
            stats[team_name]["red_cards"] += int(e.get("red_cards") or 0)
            full = clean_text(f"{e.get('first_name','')} {e.get('second_name','')}")
            player_rows.append({
                "full_name": full,
                "web_name": clean_text(e.get("web_name")),
                "goals": int(e.get("goals_scored") or 0),
                "assists": int(e.get("assists") or 0),
            })

    if relevant_finished:
        latest_event = max(int(fx.get("event") or 0) for fx in relevant_finished)
        event_info = next((e for e in bootstrap.get("events", []) if int(e.get("id", -1)) == latest_event), None) if bootstrap else None
        if event_info and event_info.get("finished"):
            state["label"] = f"GW{latest_event} – ferdig"
        else:
            state["label"] = f"GW{latest_event} – pågår"
        state["source_status"] = "LIVE"

    # Deterministic table: PL metrics first, alphabetic fallback for complete ties.
    ordered = sorted(
        teams,
        key=lambda team: (-stats[team]["points"], -stats[team]["gd"], -stats[team]["gf"], norm(team)),
    )
    table = []
    for pos, team in enumerate(ordered, 1):
        row = {"position": pos, "team": team, **stats[team]}
        table.append(row)
    state["table"] = table

    def team_leaders(metric: str, require_positive: bool = True):
        values = {team: stats[team][metric] for team in teams}
        top = max(values.values()) if values else 0
        active = bool(relevant_finished) and (top > 0 if require_positive else True)
        leaders = {team for team, val in values.items() if val == top} if active else set()
        return {"leaders": leaders, "value": top, "active": active}

    state["bonus_leaders"]["goals"] = team_leaders("gf")
    state["bonus_leaders"]["clean_sheets"] = team_leaders("clean_sheets")
    state["bonus_leaders"]["draws"] = team_leaders("drawn")
    # Cards can be active after matches even if no cards have been given; keep neutral at zero.
    state["bonus_leaders"]["yellow_cards"] = team_leaders("yellow_cards")
    state["bonus_leaders"]["red_cards"] = team_leaders("red_cards")

    def player_leaders(metric: str):
        if not relevant_finished or not player_rows:
            return {"leaders": set(), "value": 0, "active": False}
        top = max((p[metric] for p in player_rows), default=0)
        if top <= 0:
            return {"leaders": set(), "value": top, "active": False}
        leaders = set()
        for p in player_rows:
            if p[metric] == top:
                leaders.add(norm(p["full_name"]))
                leaders.add(norm(p["web_name"]))
        return {"leaders": leaders, "value": top, "active": True}

    state["bonus_leaders"]["top_scorer"] = player_leaders("goals")
    state["bonus_leaders"]["assists"] = player_leaders("assists")
    return state


def score_participant(person: dict[str, Any], live: dict[str, Any]) -> dict[str, Any]:
    current_pos = {row["team"]: row["position"] for row in live["table"]}
    detail = []
    table_points = 0
    diffs = []
    for tip in person["table_tips"]:
        team = tip["team"]
        predicted = tip["tip_place"]
        actual = current_pos.get(team)
        if actual is None:
            diff = 99
            pts = 0
        else:
            diff = abs(predicted - actual)
            pts = SCORING.get(diff, 0)
        table_points += pts
        diffs.append(diff)
        detail.append({"Lag": team, "Tippet": predicted, "Nå": actual, "Avvik": diff, "Poeng": pts})

    exact = sum(d == 0 for d in diffs)
    one = sum(d == 1 for d in diffs)
    two = sum(d == 2 for d in diffs)
    three = sum(d == 3 for d in diffs)
    over = sum(d > 3 for d in diffs)
    within1 = 100 * sum(d <= 1 for d in diffs) / len(diffs) if diffs else 0.0
    avg_diff = sum(diffs) / len(diffs) if diffs else 0.0

    bonus_rows = []
    bonus_points = 0
    for b in person["bonus_tips"]:
        key = b["key"]
        leader = live["bonus_leaders"].get(key) if key else None
        status = "neutral"
        status_text = "Ikke i gang"
        points = 0
        if leader and leader["active"]:
            if b["answer_type"].strip().casefold() == "lag":
                hit = b["pick"] in leader["leaders"]
            else:
                hit = norm(b["pick"]) in leader["leaders"]
            if hit:
                status = "good"
                status_text = "Leder" if len(leader["leaders"]) == 1 else "Delt ledelse"
                points = BONUS_POINTS
            else:
                status = "bad"
                status_text = "Ikke leder"
        bonus_points += points
        bonus_rows.append({
            "key": key,
            "label": CATEGORY_LABELS.get(key, b["category"]),
            "pick": b["pick"],
            "status": status,
            "status_text": status_text,
            "points": points,
        })

    return {
        **person,
        "table_points": table_points,
        "bonus_points": bonus_points,
        "total_points": table_points + bonus_points,
        "exact": exact,
        "one": one,
        "two": two,
        "three": three,
        "over": over,
        "within1": within1,
        "avg_diff": avg_diff,
        "table_detail": detail,
        "bonus_detail": bonus_rows,
    }


def medal(rank: int) -> str:
    return {1: "🥇", 2: "🥈", 3: "🥉"}.get(rank, f"#{rank}")


def initials(name: str) -> str:
    parts = [p for p in name.split() if p]
    return "".join(p[0].upper() for p in parts[:2]) or "?"


def participant_image(name: str) -> str | None:
    slug = slug_name(name)
    for ext in ("jpg", "jpeg", "png", "webp"):
        p = IMAGE_DIR / f"{slug}.{ext}"
        if p.exists():
            return str(p)
    return None


def bonus_html(row: dict[str, Any]) -> str:
    status = row["status"]
    icon = {"good": "●", "bad": "●", "neutral": "●"}[status]
    return f"""
    <div class="bonus-row bonus-{status}">
      <div class="bonus-copy">
        <div class="bonus-label">{html.escape(row['label'])}</div>
        <div class="bonus-pick">{html.escape(row['pick'] or '—')}</div>
      </div>
      <div class="bonus-state">{icon} {html.escape(row['status_text'])}</div>
    </div>
    """


def render_participant_card(p: dict[str, Any], rank: int):
    with st.container(border=True):
        left, right = st.columns([1.7, 1.2], gap="large")
        with left:
            photo, identity = st.columns([0.65, 2.35], vertical_alignment="center")
            with photo:
                img = participant_image(p["name"])
                if img:
                    st.image(img, use_container_width=True)
                else:
                    st.markdown(f'<div class="avatar">{html.escape(initials(p["name"]))}</div>', unsafe_allow_html=True)
            with identity:
                st.markdown(f"### {medal(rank)} {p['name']}")
                st.markdown(f'<div class="big-score">{p["total_points"]} p</div>', unsafe_allow_html=True)
                st.caption(f'Tabell {p["table_points"]} p · Bonus {p["bonus_points"]} p')

            a, b, c, d, e = st.columns(5)
            a.metric("🎯 Fulltreff", p["exact"])
            b.metric("±1 plass", p["one"])
            c.metric("±2 plasser", p["two"])
            d.metric("±3 plasser", p["three"])
            e.metric(">3 plasser", p["over"])

            x, y = st.columns(2)
            x.metric("Treffsikkerhet ±1", f'{p["within1"]:.0f}%')
            y.metric("Snittavvik", f'{p["avg_diff"]:.1f} pl.')

            with st.expander("Se alle 20 tabelltips"):
                df = pd.DataFrame(p["table_detail"])
                st.dataframe(df, hide_index=True, use_container_width=True)

        with right:
            st.markdown("#### Bonustips")
            if live_state["source_status"] == "GW0":
                st.caption("Nøytral status frem til første relevante kampdata finnes.")
            for row in p["bonus_detail"]:
                st.markdown(bonus_html(row), unsafe_allow_html=True)


def render_validation(people: list[dict[str, Any]]):
    with st.sidebar.expander("Innleveringskontroll", expanded=False):
        for p in people:
            if p["errors"]:
                st.error(f'{p["name"]}: ' + " | ".join(p["errors"]))
            elif p["warnings"]:
                st.warning(f'{p["name"]}: ' + " | ".join(p["warnings"]))
            else:
                st.success(f'{p["name"]}: OK')


st.markdown("""
<style>
.block-container {padding-top: 1.6rem; padding-bottom: 3rem;}
.big-score {font-size: 2.15rem; font-weight: 750; line-height: 1.05; margin-bottom: .25rem;}
.avatar {width: 100%; aspect-ratio: 1/1; border-radius: 18px; background: var(--secondary-background-color);
 display:flex; align-items:center; justify-content:center; font-size: 2.2rem; font-weight: 750; border: 1px solid rgba(128,128,128,.22);}
.bonus-row {display:flex; gap:.7rem; align-items:center; justify-content:space-between; padding:.58rem .72rem;
 border-radius:.7rem; margin:.38rem 0; border:1px solid rgba(128,128,128,.18);}
.bonus-copy {min-width:0;}
.bonus-label {font-size:.78rem; opacity:.72; line-height:1.15;}
.bonus-pick {font-weight:650; margin-top:.1rem; line-height:1.2; overflow-wrap:anywhere;}
.bonus-state {font-size:.78rem; font-weight:700; white-space:nowrap;}
.bonus-good {background:rgba(34,197,94,.12); border-color:rgba(34,197,94,.28);}
.bonus-good .bonus-state {color:rgb(22,163,74);}
.bonus-bad {background:rgba(239,68,68,.10); border-color:rgba(239,68,68,.24);}
.bonus-bad .bonus-state {color:rgb(220,38,38);}
.bonus-neutral {background:rgba(128,128,128,.07);}
.bonus-neutral .bonus-state {opacity:.62;}
@media (max-width: 700px) {.big-score {font-size:1.75rem;} .bonus-state {white-space:normal; text-align:right;}}
</style>
""", unsafe_allow_html=True)

st.title("🏆 PL-Tippen 2026/27")
st.caption("Last opp ferdigutfylte tippeark. Før sesongstart scores tabelltipsene mot GW0 – alfabetisk starttabell.")

uploaded_files = st.sidebar.file_uploader(
    "Last opp tippeark",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Testfiler lagres ikke permanent. Fjern dem fra opplasteren når testen er ferdig.",
)

if not uploaded_files:
    st.info("Last opp 2–5 testark i sidepanelet for å se leaderboard og deltakerkort.")
    st.markdown("**V3 inneholder:** Excel-import, GW0-scoring, deltakerkort, bonusstatus på høyre side og innleveringskontroll.")
    st.stop()

people = []
for file in uploaded_files:
    try:
        people.append(parse_tip_file(file))
    except Exception as exc:
        st.error(f"Kunne ikke lese {file.name}: {exc}")

if not people:
    st.stop()

render_validation(people)

team_sets = [set(norm(t) for t in p["teams"]) for p in people if p["teams"]]
if team_sets and any(s != team_sets[0] for s in team_sets[1:]):
    st.warning("Tippefilene inneholder ulike laglister. Bruk samme mal for alle deltakere.")

base_teams = people[0]["teams"]
live_state = build_live_state(base_teams)

if live_state.get("api_message"):
    st.sidebar.info(live_state["api_message"])

scored = [score_participant(p, live_state) for p in people]
scored.sort(key=lambda p: (-p["total_points"], p["avg_diff"], norm(p["name"])))

st.markdown(f'### {live_state["label"]}')
if live_state["source_status"] == "GW0":
    st.caption("Alle lag står på 0 poeng. Tabellen rangeres alfabetisk som en morsom startreferanse. Bonustips er nøytrale.")

# Headline leaderboard
leader_rows = []
for i, p in enumerate(scored, 1):
    leader_rows.append({
        "Plass": i,
        "Deltaker": p["name"],
        "Poeng": p["total_points"],
        "Tabell": p["table_points"],
        "Bonus": p["bonus_points"],
        "Fulltreff": p["exact"],
        "±1 %": f'{p["within1"]:.0f}%',
        "Snittavvik": round(p["avg_diff"], 1),
    })

st.dataframe(pd.DataFrame(leader_rows), hide_index=True, use_container_width=True)
st.divider()

for rank, p in enumerate(scored, 1):
    render_participant_card(p, rank)

with st.expander("Se aktuell Premier League-tabell"):
    table_df = pd.DataFrame(live_state["table"])[["position", "team", "played", "won", "drawn", "lost", "gf", "ga", "gd", "points"]]
    table_df.columns = ["#", "Lag", "K", "V", "U", "T", "MF", "MM", "+/-", "P"]
    st.dataframe(table_df, hide_index=True, use_container_width=True)
