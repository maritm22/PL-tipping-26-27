
import io
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="PL-Tippen 2026/27", page_icon="🏆", layout="wide")

SCORING = {0: 10, 1: 7, 2: 5, 3: 3}
IMAGE_DIR = Path("images")

def points_for_diff(diff):
    return SCORING.get(abs(int(diff)), 0)

def accuracy_stats(diffs):
    diffs = [abs(int(x)) for x in diffs]
    n = len(diffs)
    if not n:
        return {"exact":0,"one":0,"two":0,"three":0,"over":0,"within1":0.0,"avg":0.0}
    return {
        "exact": sum(d == 0 for d in diffs),
        "one": sum(d == 1 for d in diffs),
        "two": sum(d == 2 for d in diffs),
        "three": sum(d == 3 for d in diffs),
        "over": sum(d > 3 for d in diffs),
        "within1": 100 * sum(d <= 1 for d in diffs) / n,
        "avg": sum(diffs) / n,
    }

def participant_image(name):
    safe = "".join(c.lower() if c.isalnum() else "_" for c in name).strip("_")
    for ext in ("jpg", "jpeg", "png", "webp"):
        p = IMAGE_DIR / f"{safe}.{ext}"
        if p.exists():
            return str(p)
    return None

def participant_card(name, rank, points, diffs, bonus_points=0, gw_points=None):
    s = accuracy_stats(diffs)
    c1, c2 = st.columns([1, 3])
    with c1:
        img = participant_image(name)
        if img:
            st.image(img, use_container_width=True)
        else:
            st.markdown(
                f"<div style='height:150px;border-radius:16px;background:#eee;"
                f"display:flex;align-items:center;justify-content:center;font-size:44px'>"
                f"{name[:1].upper()}</div>", unsafe_allow_html=True
            )
    with c2:
        st.subheader(name)
        st.markdown(f"### {points} poeng · #{rank}")
        if gw_points is not None:
            st.caption(f"Denne runden: +{gw_points} poeng")

    a,b,c,d,e = st.columns(5)
    a.metric("🎯 Fulltreff", s["exact"])
    b.metric("±1 plass", s["one"])
    c.metric("±2 plasser", s["two"])
    d.metric("±3 plasser", s["three"])
    e.metric(">3 plasser", s["over"])

    x,y,z = st.columns(3)
    x.metric("Treffsikkerhet ±1", f'{s["within1"]:.0f}%')
    y.metric("Snittavvik", f'{s["avg"]:.1f} pl.')
    z.metric("Bonuspoeng", bonus_points)

    return s

def demo():
    st.title("🏆 PL-Tippen 2026/27")
    st.caption("Prototype av deltakerkort og treffsikkerhetsstatistikk")
    demo_people = [
        ("Henrik", 1, 247, [0,0,0,0,0,0,1,1,1,1,1,2,2,2,2,3,3,4,5,6], 30, 18),
        ("Martin", 2, 241, [0,0,0,0,1,1,1,1,1,1,2,2,2,2,2,3,3,3,4,5], 15, 11),
    ]
    for name, rank, pts, diffs, bonus, gw in demo_people:
        with st.container(border=True):
            participant_card(name, rank, pts, diffs, bonus, gw)

def main():
    st.sidebar.title("PL-Tippen")
    mode = st.sidebar.radio("Visning", ["Dashboard-demo", "Om statistikken"])

    if mode == "Dashboard-demo":
        demo()
    else:
        st.title("Statistikken")
        st.markdown("""
**Fulltreff** = tippet plassering er lik aktuell plassering.  
**±1 / ±2 / ±3** = absolutt avvik mellom tippet og aktuell plassering.  
**>3** = fire eller flere plasser unna.  
**Treffsikkerhet ±1** = andel av de 20 lagene som er enten eksakt eller maksimalt én plass unna.  
**Snittavvik** = gjennomsnittlig absolutt plasseringsavvik for alle 20 lag.

Poengregelen er fortsatt **10 / 7 / 5 / 3 / 0 poeng** ved henholdsvis 0 / 1 / 2 / 3 / >3 plasser avvik.
        """)

if __name__ == "__main__":
    main()
