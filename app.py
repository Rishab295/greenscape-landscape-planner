import os
import re
import time
import base64
from io import BytesIO

import requests
import pandas as pd
import streamlit as st

from PIL import Image
from dotenv import load_dotenv
from google import genai

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# ENVIRONMENT
# ============================================================

load_dotenv()


def get_secret(name):
    """Read a deployment secret from Streamlit Secrets first, then local .env."""
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name)


GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
CLOUDFLARE_API_TOKEN = get_secret("CLOUDFLARE_API_TOKEN")
CLOUDFLARE_ACCOUNT_ID = get_secret("CLOUDFLARE_ACCOUNT_ID")


# ============================================================
# MODELS
# ============================================================

# Current Gemini Flash models.
# The app will automatically try the next model if one
# temporarily returns 503 / unavailable.

GEMINI_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash-lite",
]

# Cloudflare FLUX.2 Klein 9B
CLOUDFLARE_MODEL = (
    "@cf/black-forest-labs/"
    "flux-2-klein-9b"
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Greenscape | Landscape Planning & Estimation",
    page_icon="🌱",
    layout="wide"
)



# ============================================================
# THEME
# ============================================================

st.session_state.dark_mode = True

# ============================================================
# CSS — GREENSCAPE PROFESSIONAL UI
# ============================================================

st.markdown(
    """
    <style>
/* ============================================================
   GREENSCAPE — MODERN PROFESSIONAL UI
   Light/Dark theme is controlled by body class injected below.
   ============================================================ */

:root{
    --gs-accent:#2f6b4f;
    --gs-accent-strong:#1f4b38;
    --gs-accent-soft:#e8f0eb;
    --gs-bg:#f5f6f4;
    --gs-surface:#ffffff;
    --gs-surface-2:#f9faf9;
    --gs-border:#dfe4e1;
    --gs-text:#202824;
    --gs-muted:#707b75;
    --gs-shadow:0 1px 2px rgba(24,38,30,.05);
}

html, body, [data-testid="stAppViewContainer"]{
    background:var(--gs-bg)!important;
}

.stApp{
    background:var(--gs-bg)!important;
    color:var(--gs-text)!important;
}

.block-container{
    max-width:1280px!important;
    padding:0 28px 50px!important;
}

#MainMenu, footer{visibility:hidden;}
header{background:transparent!important;}

/* ---------- Top navigation ---------- */

.gs-nav{
    height:68px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    border-bottom:1px solid var(--gs-border);
    margin-bottom:30px;
}

.gs-brand{
    display:flex;
    align-items:center;
    gap:11px;
}

.gs-brand-mark{
    width:34px;
    height:34px;
    border-radius:9px;
    background:var(--gs-accent-strong);
    color:white;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:12px;
    font-weight:800;
    letter-spacing:-.4px;
}

.gs-brand-name{
    font-size:20px;
    font-weight:800;
    letter-spacing:-.5px;
    color:var(--gs-text);
}

.gs-brand-sub{
    color:var(--gs-muted);
    font-size:10px;
    margin-top:2px;
}

.gs-nav-right{
    display:flex;
    align-items:center;
    gap:18px;
    color:var(--gs-muted);
    font-size:11px;
}

/* ---------- Page heading ---------- */

.gs-heading{
    margin-bottom:24px;
}

.gs-heading-row{
    display:flex;
    align-items:flex-end;
    justify-content:space-between;
    gap:20px;
}

.gs-heading h1{
    margin:0;
    color:var(--gs-text);
    font-size:30px;
    line-height:1.1;
    font-weight:780;
    letter-spacing:-1px;
}

.gs-heading p{
    margin:7px 0 0;
    color:var(--gs-muted);
    font-size:12px;
}

.gs-status{
    padding:7px 10px;
    border:1px solid var(--gs-border);
    background:var(--gs-surface);
    border-radius:6px;
    color:var(--gs-muted);
    font-size:10px;
    white-space:nowrap;
}

/* ---------- Section blocks ---------- */

.gs-section-head{
    display:flex;
    align-items:center;
    justify-content:space-between;
    margin:30px 0 11px;
}

.gs-section-title{
    color:var(--gs-text);
    font-size:16px;
    font-weight:760;
    letter-spacing:-.25px;
}

.gs-section-desc{
    color:var(--gs-muted);
    font-size:10px;
}

.gs-card{
    background:var(--gs-surface);
    border:1px solid var(--gs-border);
    border-radius:8px;
    padding:16px 17px;
    box-shadow:var(--gs-shadow);
}

.gs-card-title{
    color:var(--gs-text);
    font-size:12px;
    font-weight:750;
    margin-bottom:10px;
}

.gs-mini-label{
    color:var(--gs-muted);
    font-size:9px;
    text-transform:uppercase;
    letter-spacing:.8px;
    font-weight:800;
    margin-bottom:5px;
}

/* ---------- Inputs ---------- */

label,
.stSelectbox label,
.stTextInput label,
.stNumberInput label,
.stTextArea label,
.stMultiSelect label,
.stFileUploader label{
    color:var(--gs-text)!important;
    font-size:11px!important;
    font-weight:650!important;
}

div[data-baseweb="input"]>div,
div[data-baseweb="select"]>div,
div[data-baseweb="textarea"]>div{
    background:var(--gs-surface)!important;
    border:1px solid var(--gs-border)!important;
    border-radius:6px!important;
    box-shadow:none!important;
    min-height:38px;
}

div[data-baseweb="input"] input,
div[data-baseweb="textarea"] textarea,
div[data-baseweb="select"] input,
div[data-baseweb="select"] span{
    color:var(--gs-text)!important;
    -webkit-text-fill-color:var(--gs-text)!important;
}

div[data-baseweb="input"]>div:focus-within,
div[data-baseweb="select"]>div:focus-within,
div[data-baseweb="textarea"]>div:focus-within{
    border-color:var(--gs-accent)!important;
    box-shadow:0 0 0 1px var(--gs-accent)!important;
}

/* ---------- Buttons ---------- */

.stButton>button{
    border-radius:6px!important;
    min-height:38px;
    font-size:11px;
    font-weight:700;
    box-shadow:none!important;
}

.stButton>button[kind="primary"]{
    background:var(--gs-accent-strong)!important;
    border:1px solid var(--gs-accent-strong)!important;
    color:#fff!important;
}

.stButton>button[kind="primary"]:hover{
    background:var(--gs-accent)!important;
    border-color:var(--gs-accent)!important;
}

.stButton>button:not([kind="primary"]){
    background:var(--gs-surface)!important;
    border:1px solid var(--gs-border)!important;
    color:var(--gs-text)!important;
}

/* ---------- Alerts / tables / expanders ---------- */

div[data-testid="stAlert"]{
    border-radius:6px!important;
    box-shadow:none!important;
    font-size:11px!important;
}

section[data-testid="stFileUploaderDropzone"]{
    background:var(--gs-surface)!important;
    border:1px dashed #b8c4bd!important;
    border-radius:6px!important;
}

div[data-testid="stDataFrame"]{
    border:1px solid var(--gs-border);
    border-radius:7px;
    overflow:hidden;
}

details{
    background:var(--gs-surface)!important;
    border:1px solid var(--gs-border)!important;
    border-radius:7px!important;
    box-shadow:none!important;
}

div[data-testid="stMetric"]{
    background:var(--gs-surface);
    border:1px solid var(--gs-border);
    border-radius:7px;
    box-shadow:none;
    padding:10px 12px;
}

div[data-testid="stMetricLabel"]{
    color:var(--gs-muted);
    font-size:10px;
}

div[data-testid="stMetricValue"]{
    color:var(--gs-accent-strong);
    font-weight:780;
}

/* ---------- Scope ---------- */

.gs-scope-note{
    background:var(--gs-accent-soft);
    border:1px solid #d3dfd8;
    border-radius:6px;
    padding:10px 12px;
    color:var(--gs-accent-strong);
    font-size:11px;
    margin-bottom:14px;
}

.gs-result{
    background:var(--gs-accent-strong);
    color:white;
    border-radius:6px;
    padding:10px 13px;
    margin:22px 0 11px;
    font-size:12px;
    font-weight:750;
}

/* ---------- Sidebar ---------- */

[data-testid="stSidebar"]{
    background:var(--gs-surface)!important;
    border-right:1px solid var(--gs-border);
}

.gs-side{
    padding:4px 2px;
}

.gs-side-brand{
    padding:5px 5px 18px;
    border-bottom:1px solid var(--gs-border);
}

.gs-side-name{
    color:var(--gs-text);
    font-size:19px;
    font-weight:800;
    letter-spacing:-.4px;
}

.gs-side-sub{
    color:var(--gs-muted);
    font-size:10px;
    margin-top:3px;
}

.gs-side-label{
    color:var(--gs-muted);
    font-size:9px;
    text-transform:uppercase;
    letter-spacing:1px;
    font-weight:800;
    margin:19px 5px 7px;
}

.gs-side-row{
    color:var(--gs-text);
    font-size:11px;
    padding:7px 8px;
    border-radius:5px;
    margin-bottom:2px;
}

.gs-side-row.active{
    background:var(--gs-accent-soft);
    color:var(--gs-accent-strong);
    font-weight:750;
}

.gs-footer{
    border-top:1px solid var(--gs-border);
    margin-top:42px;
    padding-top:12px;
    color:var(--gs-muted);
    font-size:9px;
}

/* ---------- Dark theme ---------- */

body.gs-dark{
    --gs-accent:#6ea88a;
    --gs-accent-strong:#77b293;
    --gs-accent-soft:#20352b;
    --gs-bg:#101412;
    --gs-surface:#171c19;
    --gs-surface-2:#1c221f;
    --gs-border:#2d3530;
    --gs-text:#edf2ef;
    --gs-muted:#9ca8a1;
    --gs-shadow:0 1px 2px rgba(0,0,0,.25);
}

body.gs-dark .stApp,
body.gs-dark [data-testid="stAppViewContainer"]{
    background:var(--gs-bg)!important;
}

body.gs-dark [data-testid="stSidebar"]{
    background:var(--gs-surface)!important;
}

body.gs-dark div[data-baseweb="input"]>div,
body.gs-dark div[data-baseweb="select"]>div,
body.gs-dark div[data-baseweb="textarea"]>div,
body.gs-dark .stButton>button:not([kind="primary"]),
body.gs-dark details,
body.gs-dark div[data-testid="stMetric"],
body.gs-dark section[data-testid="stFileUploaderDropzone"]{
    background:var(--gs-surface)!important;
}

body.gs-dark .stButton>button:not([kind="primary"]){
    color:var(--gs-text)!important;
}

body.gs-dark div[data-baseweb="input"] input,
body.gs-dark div[data-baseweb="textarea"] textarea,
body.gs-dark div[data-baseweb="select"] input,
body.gs-dark div[data-baseweb="select"] span{
    color:var(--gs-text)!important;
    -webkit-text-fill-color:var(--gs-text)!important;
}

body.gs-dark .gs-scope-note{
    border-color:#31503f;
}

@media(max-width:900px){
    .gs-workflow{grid-template-columns:repeat(3,1fr)}
}

/* ============================================================
   BUTTON COLOUR FIX
   ============================================================ */

.stButton > button,
div[data-testid="stDownloadButton"] > button{
    background:#ffffff!important;
    color:#1f4938!important;
    border:1px solid #2f6b4f!important;
    border-radius:6px!important;
    box-shadow:none!important;
    font-size:11px!important;
    font-weight:700!important;
    min-height:38px!important;
}

.stButton > button:hover,
div[data-testid="stDownloadButton"] > button:hover{
    background:#e8f0eb!important;
    color:#17362a!important;
    border-color:#1f4938!important;
}

.stButton > button[kind="primary"],
.stButton > button[data-testid="baseButton-primary"],
div[data-testid="stDownloadButton"] > button[kind="primary"]{
    background:#1f5a43!important;
    color:#ffffff!important;
    border:1px solid #1f5a43!important;
}

.stButton > button[kind="primary"]:hover,
.stButton > button[data-testid="baseButton-primary"]:hover,
div[data-testid="stDownloadButton"] > button[kind="primary"]:hover{
    background:#164633!important;
    color:#ffffff!important;
    border-color:#164633!important;
}

/* Theme controls */
button[kind="primary"]{
    background:#1f5a43!important;
    color:#ffffff!important;
}



/* ============================================================
   FINAL GREENSCAPE BUTTONS
   ============================================================ */

.stButton > button,
div[data-testid="stDownloadButton"] > button{
    min-height:38px!important;
    border-radius:6px!important;
    font-size:11px!important;
    font-weight:700!important;
    box-shadow:none!important;
    background:#ffffff!important;
    color:#1f4938!important;
    border:1px solid #2f6b4f!important;
}

.stButton > button:hover,
div[data-testid="stDownloadButton"] > button:hover{
    background:#e8f0eb!important;
    color:#17362a!important;
    border-color:#1f4938!important;
}

.stButton > button[kind="primary"],
.stButton > button[data-testid="baseButton-primary"]{
    background:#1f5a43!important;
    color:#ffffff!important;
    border:1px solid #1f5a43!important;
}

.stButton > button[kind="primary"]:hover,
.stButton > button[data-testid="baseButton-primary"]:hover{
    background:#164633!important;
    color:#ffffff!important;
    border-color:#164633!important;
}


    .gs-header-actions{
        display:flex;
        align-items:center;
        gap:8px;
    }

    .gs-dark-badge{
        display:inline-flex;
        align-items:center;
        min-height:28px;
        padding:0 9px;
        border-radius:5px;
        background:#20392c;
        color:#8fc2a4;
        border:1px solid #355441;
        font-size:9px;
        font-weight:800;
        letter-spacing:.5px;
    }


        /* =====================================================
           DARK MODE — GLOBAL TEXT & CONTROL CONTRAST
           ===================================================== */

        html, body,
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"],
        [data-testid="stSidebar"],
        [data-testid="stSidebarContent"]{
            background:#0f1411!important;
            color:#edf3ef!important;
        }

        [data-testid="stHeader"]{
            background:#0f1411!important;
        }

        /* All normal Streamlit text */
        [data-testid="stMarkdownContainer"],
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        [data-testid="stMarkdownContainer"] span,
        [data-testid="stMarkdownContainer"] strong,
        [data-testid="stMarkdownContainer"] em{
            color:#dce6e0!important;
        }

        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stMarkdownContainer"] h5,
        [data-testid="stMarkdownContainer"] h6{
            color:#f2f6f4!important;
        }

        h1,h2,h3,h4,h5,h6{
            color:#f2f6f4!important;
        }

        p,li{
            color:#dce6e0!important;
        }

        small,
        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] *{
            color:#9eaaa3!important;
        }

        /* Section headings generated by Streamlit */
        .gs-section-title,
        .gs-section-desc,
        .gs-card-title,
        .gs-mini-label,
        .gs-heading h1,
        .gs-heading p,
        .gs-brand-name,
        .gs-brand-sub,
        .gs-nav-right,
        .gs-status,
        .gs-side-name,
        .gs-side-sub,
        .gs-side-label,
        .gs-side-row{
            opacity:1!important;
        }

        .gs-section-title,
        .gs-card-title,
        .gs-side-name,
        .gs-brand-name,
        .gs-heading h1{
            color:#f2f6f4!important;
        }

        .gs-section-desc,
        .gs-heading p,
        .gs-brand-sub,
        .gs-nav-right,
        .gs-side-sub,
        .gs-side-label{
            color:#9eaaa3!important;
        }

        .gs-side-row{
            color:#cbd7d0!important;
        }

        .gs-side-row.active{
            color:#a7d2b9!important;
            background:#1d3528!important;
        }

        /* Labels */
        label,
        .stSelectbox label,
        .stTextInput label,
        .stNumberInput label,
        .stTextArea label,
        .stMultiSelect label,
        .stFileUploader label,
        [data-testid="stWidgetLabel"]{
            color:#dce6e0!important;
        }

        /* Text fields */
        div[data-baseweb="input"]>div,
        div[data-baseweb="select"]>div,
        div[data-baseweb="textarea"]>div{
            background:#1a201d!important;
            border:1px solid #38443d!important;
            color:#f0f5f2!important;
        }

        div[data-baseweb="input"] input,
        div[data-baseweb="textarea"] textarea,
        div[data-baseweb="select"] input{
            color:#f0f5f2!important;
            -webkit-text-fill-color:#f0f5f2!important;
            caret-color:#9bd1b0!important;
        }

        div[data-baseweb="select"] span{
            color:#f0f5f2!important;
            -webkit-text-fill-color:#f0f5f2!important;
        }

        input::placeholder,
        textarea::placeholder{
            color:#8c9991!important;
            -webkit-text-fill-color:#8c9991!important;
            opacity:1!important;
        }

        /* Dropdown menus */
        [data-baseweb="popover"],
        [data-baseweb="menu"],
        [role="listbox"]{
            background:#1a201d!important;
            color:#edf3ef!important;
            border-color:#39463f!important;
        }

        [role="option"]{
            background:#1a201d!important;
            color:#edf3ef!important;
        }

        [role="option"]:hover,
        [aria-selected="true"]{
            background:#244233!important;
            color:#ffffff!important;
        }

        /* Multiselect chips */
        [data-baseweb="tag"]{
            background:#28513d!important;
            color:#ffffff!important;
        }

        [data-baseweb="tag"] span{
            color:#ffffff!important;
        }

        /* Number input controls */
        button[aria-label="Increment"],
        button[aria-label="Decrement"]{
            color:#dce6e0!important;
            background:#1a201d!important;
        }

        /* Buttons */
        .stButton>button,
        div[data-testid="stDownloadButton"]>button{
            background:#1a201d!important;
            color:#e8f2ec!important;
            border:1px solid #4a6657!important;
        }

        .stButton>button:hover,
        div[data-testid="stDownloadButton"]>button:hover{
            background:#244233!important;
            color:#ffffff!important;
            border-color:#5e8c70!important;
        }

        .stButton>button[kind="primary"],
        .stButton>button[data-testid="baseButton-primary"]{
            background:#2f7655!important;
            color:#ffffff!important;
            border-color:#2f7655!important;
        }

        .stButton>button[kind="primary"]:hover,
        .stButton>button[data-testid="baseButton-primary"]:hover{
            background:#3b8d67!important;
            color:#ffffff!important;
            border-color:#3b8d67!important;
        }

        /* File uploader */
        section[data-testid="stFileUploaderDropzone"]{
            background:#171d1a!important;
            border:1px dashed #4a5d51!important;
        }

        section[data-testid="stFileUploaderDropzone"] *,
        section[data-testid="stFileUploaderDropzone"] small{
            color:#cbd7d0!important;
        }

        /* Alerts */
        div[data-testid="stAlert"]{
            color:#e8f0eb!important;
        }

        div[data-testid="stAlert"] p,
        div[data-testid="stAlert"] span{
            color:inherit!important;
        }

        /* Expanders */
        details,
        [data-testid="stExpander"]{
            background:#171d1a!important;
            border-color:#303b35!important;
        }

        details summary,
        [data-testid="stExpander"] summary{
            color:#e3ece7!important;
        }

        /* Metrics */
        div[data-testid="stMetric"]{
            background:#171d1a!important;
            border-color:#303b35!important;
        }

        div[data-testid="stMetricLabel"]{
            color:#9eaaa3!important;
        }

        div[data-testid="stMetricValue"],
        div[data-testid="stMetricDelta"]{
            color:#f0f5f2!important;
        }

        /* Dataframe */
        div[data-testid="stDataFrame"]{
            border-color:#303b35!important;
        }

        /* Dark-mode badge */
        .gs-dark-badge{
            background:#20392c!important;
            color:#a7d2b9!important;
            border-color:#3b624d!important;
        }

        /* Footer */
        .gs-footer,
        .gs-footer *{
            color:#7f8c84!important;
        }

        /* Hide the native sidebar collapse affordance if it is visually
           conflicting with the custom dark sidebar. */
        [data-testid="stSidebarCollapseButton"] button{
            color:#aab7af!important;
        }


        [data-testid="stSidebar"]{
            background:#111613!important;
        }

        [data-testid="stSidebarContent"]{
            background:#111613!important;
        }

        .gs-side{
            background:#111613!important;
            min-height:100vh;
        }

        .gs-side-brand{
            background:#171d1a!important;
            border:1px solid #2d3832!important;
            border-radius:5px!important;
            padding:12px 10px 13px!important;
        }


        /* Sidebar removed — full width workspace */
        [data-testid="stSidebar"],
        [data-testid="stSidebarContent"],
        [data-testid="stSidebarNav"]{
            display:none!important;
            width:0!important;
            min-width:0!important;
        }

        section[data-testid="stSidebar"]{
            display:none!important;
        }

        [data-testid="stSidebarCollapseButton"]{
            display:none!important;
        }


        /* =====================================================
           FULL-WIDTH TOP HEADER
           ===================================================== */

        .block-container{
            max-width:1380px!important;
            padding:0 42px 56px!important;
        }

        .gs-topbar{
            width:100%;
            min-height:72px;
            display:flex;
            align-items:center;
            justify-content:space-between;
            box-sizing:border-box;
            border-bottom:1px solid #29332e;
            margin:0 0 34px 0;
            padding:0 2px;
        }

        .gs-topbar-brand{
            display:flex;
            align-items:center;
            gap:11px;
        }

        .gs-brand-mark{
            width:38px!important;
            height:38px!important;
            min-width:38px!important;
            border-radius:7px!important;
            background:#2b7353!important;
            color:#ffffff!important;
            display:flex!important;
            align-items:center!important;
            justify-content:center!important;
            font-size:11px!important;
            font-weight:800!important;
            letter-spacing:.2px!important;
        }

        .gs-brand-copy{
            display:flex;
            flex-direction:column;
            justify-content:center;
        }

        .gs-brand-name{
            color:#f1f5f3!important;
            font-size:19px!important;
            line-height:1.1!important;
            font-weight:800!important;
            letter-spacing:-.3px!important;
            margin:0!important;
        }

        .gs-brand-sub{
            color:#8e9b94!important;
            font-size:9px!important;
            line-height:1.2!important;
            margin-top:4px!important;
        }

        .gs-topbar-right{
            display:flex;
            align-items:center;
            gap:11px;
        }

        .gs-workspace-label{
            color:#8e9b94!important;
            font-size:10px!important;
            white-space:nowrap;
        }

        .gs-topbar-divider{
            width:1px;
            height:18px;
            background:#303b35;
            display:block;
        }

        .gs-dark-badge{
            display:inline-flex!important;
            align-items:center!important;
            justify-content:center!important;
            min-height:27px!important;
            padding:0 9px!important;
            border-radius:5px!important;
            background:#1d3528!important;
            color:#a8d1b9!important;
            border:1px solid #3a624d!important;
            font-size:8px!important;
            font-weight:800!important;
            letter-spacing:.65px!important;
        }

        .gs-page-heading{
            display:flex;
            align-items:flex-end;
            justify-content:space-between;
            gap:20px;
            margin:0 0 31px 0;
        }

        .gs-page-heading h1{
            margin:0!important;
            color:#f1f5f3!important;
            font-size:30px!important;
            line-height:1.15!important;
            font-weight:780!important;
            letter-spacing:-.8px!important;
        }

        .gs-page-heading p{
            margin:7px 0 0!important;
            color:#8f9c95!important;
            font-size:11px!important;
        }

        .gs-page-heading .gs-status{
            display:inline-flex!important;
            align-items:center!important;
            min-height:28px!important;
            padding:0 10px!important;
            border-radius:5px!important;
            background:#171d1a!important;
            color:#9ca9a1!important;
            border:1px solid #303b35!important;
            font-size:9px!important;
            white-space:nowrap!important;
        }

        /* Remove any old header spacing/classes that could overlap. */
        .gs-nav,
        .gs-heading,
        .gs-heading-row{
            display:none!important;
        }


        [data-testid="stAppViewContainer"]{
            margin-left:0!important;
        }

        [data-testid="stMain"]{
            margin-left:0!important;
            width:100%!important;
        }

        [data-testid="stMainBlockContainer"]{
            max-width:1380px!important;
        }


        /* FINAL HEADER REPAIR */
        .gs-topbar{
            position:relative!important;
            z-index:5!important;
            width:100%!important;
            min-height:72px!important;
            display:flex!important;
            align-items:center!important;
            justify-content:space-between!important;
            gap:30px!important;
            box-sizing:border-box!important;
        }

        .gs-topbar-brand,
        .gs-topbar-right{
            display:flex!important;
            align-items:center!important;
        }

        .gs-topbar-right{
            margin-left:auto!important;
            flex-shrink:0!important;
        }

        .gs-dark-badge{
            white-space:nowrap!important;
        }

        .gs-workspace-label{
            white-space:nowrap!important;
        }

        .gs-page-heading{
            position:relative!important;
            z-index:4!important;
        }


        /* =====================================================
           FINAL HEADER CLIPPING FIX
           ===================================================== */

        .gs-topbar{
            height:auto!important;
            min-height:82px!important;
            padding:10px 0!important;
            margin:0 0 34px 0!important;
            overflow:visible!important;
            align-items:center!important;
        }

        .gs-topbar-brand{
            height:56px!important;
            min-height:56px!important;
            display:flex!important;
            align-items:center!important;
            overflow:visible!important;
        }

        .gs-brand-mark{
            flex:0 0 38px!important;
            width:38px!important;
            height:38px!important;
            margin:0!important;
        }

        .gs-brand-copy{
            height:52px!important;
            min-height:52px!important;
            display:flex!important;
            flex-direction:column!important;
            justify-content:center!important;
            align-items:flex-start!important;
            overflow:visible!important;
            padding:2px 0!important;
            box-sizing:border-box!important;
        }

        .gs-brand-name{
            display:block!important;
            height:27px!important;
            min-height:27px!important;
            line-height:27px!important;
            font-size:20px!important;
            font-weight:800!important;
            margin:0!important;
            padding:0!important;
            overflow:visible!important;
            white-space:nowrap!important;
            color:#f1f5f3!important;
            transform:none!important;
        }

        .gs-brand-sub{
            display:block!important;
            height:15px!important;
            min-height:15px!important;
            line-height:15px!important;
            font-size:9px!important;
            margin:1px 0 0!important;
            padding:0!important;
            overflow:visible!important;
            white-space:nowrap!important;
            color:#8e9b94!important;
        }

        .gs-topbar-right{
            height:52px!important;
            min-height:52px!important;
            align-items:center!important;
            overflow:visible!important;
        }

        .gs-dark-badge,
        .gs-workspace-label{
            line-height:normal!important;
        }

        .gs-page-heading{
            margin-top:0!important;
            padding-top:0!important;
            overflow:visible!important;
        }

        /* Prevent Streamlit's markdown wrapper from clipping
           the first line of custom header HTML. */
        .gs-topbar,
        .gs-topbar *{
            overflow:visible!important;
        }


        /* =====================================================
           FINAL HEADER CLEANUP
           ===================================================== */

        /* Give the header a real top breathing space so the
           Greenscape wordmark can never touch/cross the top edge. */
        .gs-topbar{
            min-height:96px!important;
            height:96px!important;
            padding:18px 0 16px!important;
            margin:0 0 34px!important;
            box-sizing:border-box!important;
            display:flex!important;
            align-items:center!important;
            overflow:visible!important;
        }

        .gs-topbar-brand{
            height:60px!important;
            min-height:60px!important;
            display:flex!important;
            align-items:center!important;
            overflow:visible!important;
        }

        .gs-brand-copy{
            height:58px!important;
            min-height:58px!important;
            padding:3px 0!important;
            display:flex!important;
            justify-content:center!important;
            align-items:flex-start!important;
            overflow:visible!important;
            box-sizing:border-box!important;
        }

        .gs-brand-name{
            display:block!important;
            height:29px!important;
            min-height:29px!important;
            line-height:29px!important;
            margin:0!important;
            padding:0!important;
            font-size:20px!important;
            font-weight:800!important;
            white-space:nowrap!important;
            overflow:visible!important;
            color:#f1f5f3!important;
        }

        .gs-brand-sub{
            display:block!important;
            height:16px!important;
            min-height:16px!important;
            line-height:16px!important;
            margin:2px 0 0!important;
            padding:0!important;
            font-size:9px!important;
            white-space:nowrap!important;
            overflow:visible!important;
            color:#8e9b94!important;
        }

        .gs-topbar-right{
            height:58px!important;
            min-height:58px!important;
            margin-left:auto!important;
            display:flex!important;
            align-items:center!important;
            overflow:visible!important;
        }

        /* The dark theme is fixed; no theme control is shown. */
        .gs-dark-badge,
        .gs-topbar-divider{
            display:none!important;
        }

        .gs-workspace-label{
            color:#8e9b94!important;
            font-size:10px!important;
            white-space:nowrap!important;
        }

        .gs-page-heading{
            margin-top:0!important;
            padding-top:0!important;
            overflow:visible!important;
        }

        .gs-topbar,
        .gs-topbar *{
            overflow:visible!important;
        }


        /* Browser/Streamlit top breathing room */
        .block-container{
            padding-top:28px!important;
        }

        [data-testid="stMainBlockContainer"]{
            padding-top:0!important;
        }

</style>

    """,
    unsafe_allow_html=True
)





# ============================================================

# ============================================================
# ============================================================
# ACTIVE THEME OVERRIDES
# ============================================================

if True:
    st.markdown(
        """
        <style>
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stHeader"]{
            background:#101412!important;
        }

        .gs-nav,
        .gs-step,
        .gs-card,
        .gs-status,
        .gs-side-row,
        .gs-side-brand,
        div[data-baseweb="input"]>div,
        div[data-baseweb="select"]>div,
        div[data-baseweb="textarea"]>div,
        details,
        div[data-testid="stMetric"],
        section[data-testid="stFileUploaderDropzone"]{
            background:#171c19!important;
            border-color:#2d3530!important;
        }

        .gs-brand-name,
        .gs-heading h1,
        .gs-section-title,
        .gs-step-text,
        .gs-side-name,
        .gs-card-title,
        .gs-mini-label,
        label,
        .stSelectbox label,
        .stTextInput label,
        .stNumberInput label,
        .stTextArea label,
        .stMultiSelect label,
        .stFileUploader label{
            color:#edf2ef!important;
        }

        .gs-brand-sub,
        .gs-product,
        .gs-nav-right,
        .gs-heading p,
        .gs-status,
        .gs-section-desc,
        .gs-note,
        .gs-side-sub,
        .gs-side-row,
        .gs-footer,
        div[data-testid="stMetricLabel"]{
            color:#a5b0aa!important;
        }

        div[data-baseweb="input"] input,
        div[data-baseweb="textarea"] textarea,
        div[data-baseweb="select"] input,
        div[data-baseweb="select"] span{
            color:#edf2ef!important;
            -webkit-text-fill-color:#edf2ef!important;
        }

        .stButton > button:not([kind="primary"]),
        div[data-testid="stDownloadButton"] > button:not([kind="primary"]){
            background:#171c19!important;
            color:#78b594!important;
            border-color:#4d735f!important;
        }

        .stButton > button:not([kind="primary"]):hover{
            background:#20352b!important;
            color:#a9d2bb!important;
        }

        .stButton > button[kind="primary"],
        .stButton > button[data-testid="baseButton-primary"]{
            background:#2f7655!important;
            color:#ffffff!important;
            border-color:#2f7655!important;
        }

        .gs-scope-note{
            background:#20352b!important;
            border-color:#31503f!important;
            color:#cfe1d7!important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# ============================================================
# ============================================================
# ============================================================
# GREENSCAPE HEADER — CLEAN
# ============================================================

st.markdown(
    """<div class="gs-topbar"><div class="gs-topbar-brand"><div class="gs-brand-mark">GS</div><div class="gs-brand-copy"><div class="gs-brand-name">Greenscape</div><div class="gs-brand-sub">Landscape Planning &amp; Estimation</div></div></div><div class="gs-topbar-right"><span class="gs-workspace-label">Project workspace</span></div></div><div class="gs-page-heading"><div><h1>Project Planning</h1><p>Site information, scope selection, planting schedule, estimate and visualisation.</p></div><span class="gs-status">Draft project</span></div>""",
    unsafe_allow_html=True
)

# CHECK GEMINI KEY
# ============================================================

if not GEMINI_API_KEY:

    st.error(
        "GEMINI_API_KEY is missing from your .env file."
    )

    st.stop()


# ============================================================
# LOAD PLANT DATABASE
# ============================================================

PLANT_FILE = "plants.csv"

if not os.path.exists(PLANT_FILE):

    st.error(
        "plants.csv was not found. "
        "Keep plants.csv in the same folder as app.py."
    )

    st.stop()


try:

    plants_df = pd.read_csv(
        PLANT_FILE
    )

except Exception as e:

    st.error(
        f"Could not read plants.csv: {e}"
    )

    st.stop()


# ============================================================
# VALIDATE PLANT CSV
# ============================================================

required_columns = [
    "plant",
    "price"
]

for column in required_columns:

    if column not in plants_df.columns:

        st.error(
            f"plants.csv is missing the required "
            f"column: {column}"
        )

        st.stop()


# ============================================================
# GEMINI CLIENT
# ============================================================

try:

    gemini_client = genai.Client(
        api_key=GEMINI_API_KEY
    )

except Exception as e:

    st.error(
        f"Gemini initialization failed: {e}"
    )

    st.stop()


# ============================================================
# SESSION STATE
# ============================================================

if "landscape_result" not in st.session_state:
    st.session_state.landscape_result = None

if "generated_image" not in st.session_state:
    st.session_state.generated_image = None

if "plant_df" not in st.session_state:
    st.session_state.plant_df = pd.DataFrame()

if "budget_df" not in st.session_state:
    st.session_state.budget_df = pd.DataFrame()


# ============================================================
# CLEAN TEXT
# ============================================================

def clean_text(text):

    if text is None:
        return ""

    text = str(text)

    text = text.replace(
        "```markdown",
        ""
    )

    text = text.replace(
        "```",
        ""
    )

    return text.strip()


# ============================================================
# CREATE LANDSCAPE PROMPT
# ============================================================

def create_landscape_prompt(
    project_name,
    location,
    area,
    sunlight,
    soil,
    water,
    maintenance,
    requirements,
    additional_requirements
):

    plant_database = plants_df.to_csv(
        index=False
    )

    prompt = f"""
You are a professional landscape architect
and horticulture consultant.

Prepare a practical landscape planning proposal
for this project.

PROJECT DETAILS
===============

Project Name:
{project_name}

Location:
{location}

Area:
{area} sq.ft


SITE CONDITIONS
===============

Sunlight:
{sunlight}

Soil:
{soil}

Water Availability:
{water}

Maintenance:
{maintenance}


LANDSCAPE REQUIREMENTS
======================

{requirements}


ADDITIONAL REQUIREMENTS
=======================

{additional_requirements}


PLANT DATABASE
==============

{plant_database}


STRICT PLANT RULES
==================

1. You MUST recommend plants ONLY from plants.csv.

2. Do NOT invent plant names.

3. Use the exact plant names from the database.

4. Use the price from the database.

5. Calculate realistic quantities.

6. Lawn Grass may be measured in sq.ft.

7. Do not recommend a plant that is not present
   in the supplied database.

8. Recommend plants ONLY when a plant-related requirement
   is selected by the user. If no plant-related requirement
   is selected, the RECOMMENDED PLANTS table must be empty.

9. Do not add plants merely because they are common in
   landscape designs; follow the user's selected scope.


OUTPUT FORMAT
=============


# SITE ANALYSIS

Explain:

- Existing site condition
- Sunlight
- Soil
- Water availability
- Maintenance level
- Important observations


# LANDSCAPE CONCEPT

Explain:

- Overall landscape concept
- Design style
- Functional zones
- Visual theme


# RECOMMENDED PLANTS

Create this exact table:

| Plant | Quantity | Unit | Rate (₹) | Amount (₹) | Reason |

Only use plants from plants.csv.


# SUNLIGHT PLANNING

Explain how plants should be positioned
according to their sunlight requirements.


# IRRIGATION RECOMMENDATION

Give a practical irrigation recommendation.


# SUGGESTED LAYOUT

Explain:

- Entrance
- Lawn
- Trees
- Boundary
- Flowering plants
- Ornamental plants
- Pathways
- Special features


# MAINTENANCE PLAN

Explain:

- Weekly work
- Monthly work
- Seasonal work
- Pruning
- Fertilization
- Irrigation
- Pest monitoring


# PRELIMINARY BUDGET

Create a budget table ONLY for the selected landscape requirements above.

| Category | Estimated Cost (₹) |

STRICT BUDGET RULES:

1. Include ONLY categories that correspond to selected requirements.
2. If a feature/service was NOT selected, do NOT include its cost.
3. Do NOT automatically add soil improvement, irrigation, pathway, labour,
   hardscape, lighting, water features or contingency.
4. Plant Material & Lawn Grass may be included only when at least one
   plant-related requirement is selected and plants are recommended.
5. Do not create a Total row; the application calculates the final total.
6. Do not invent quotation categories outside the selected requirements.


# SITE VERIFICATION

List everything that should be physically
verified before execution.

Do not claim that exact measurements were taken
unless they were provided.


# CUSTOMER-FRIENDLY SUMMARY

Write a short professional summary
suitable for sending to a customer.

Clearly state that the proposal is preliminary
and final quantities should be verified at site.
"""

    return prompt


# ============================================================
# GEMINI LANDSCAPE GENERATION
# ============================================================

def generate_landscape_plan(
    project_name,
    location,
    area,
    sunlight,
    soil,
    water,
    maintenance,
    requirements,
    additional_requirements,
    uploaded_image
):

    prompt = create_landscape_prompt(
        project_name,
        location,
        area,
        sunlight,
        soil,
        water,
        maintenance,
        requirements,
        additional_requirements
    )

    last_error = None

    # ========================================================
    # TRY EACH MODEL
    # ========================================================

    for model_name in GEMINI_MODELS:

        for attempt in range(3):

            try:

                image = Image.open(
                    uploaded_image
                )

                response = (
                    gemini_client
                    .models
                    .generate_content(
                        model=model_name,
                        contents=[
                            prompt,
                            image
                        ]
                    )
                )

                if not response.text:

                    raise Exception(
                        "Gemini returned an empty response."
                    )

                return clean_text(
                    response.text
                )

            except Exception as e:

                error_text = str(e)

                last_error = (
                    f"{model_name}: "
                    f"{error_text}"
                )

                # =================================================
                # TEMPORARY 503
                # =================================================

                if (
                    "503" in error_text
                    or
                    "UNAVAILABLE" in error_text
                    or
                    "high demand"
                    in error_text.lower()
                ):

                    # Exponential backoff:
                    # 2 sec -> 4 sec -> 8 sec

                    wait_time = (
                        2 ** attempt
                    )

                    time.sleep(
                        wait_time
                    )

                    continue

                # =================================================
                # QUOTA / RATE LIMIT
                # =================================================

                if (
                    "429" in error_text
                    or
                    "RESOURCE_EXHAUSTED"
                    in error_text
                ):

                    # Try next Gemini model
                    break

                # =================================================
                # OTHER ERROR
                # =================================================

                raise Exception(
                    f"Gemini error using "
                    f"{model_name}:\n\n"
                    f"{error_text}"
                )

    # ========================================================
    # ALL MODELS FAILED
    # ========================================================

    raise Exception(
        "All Gemini models were temporarily "
        "unavailable.\n\n"
        f"Last error:\n{last_error}"
    )


# ============================================================
# EXTRACT PLANT BOQ
# ============================================================

def extract_plant_boq(
    landscape_result
):

    rows = []

    valid_plants = (
        plants_df["plant"]
        .astype(str)
        .str.strip()
        .tolist()
    )

    for line in landscape_result.splitlines():

        line = line.strip()

        if not line.startswith("|"):
            continue

        parts = [
            p.strip()
            for p in line.strip("|").split("|")
        ]

        if len(parts) < 5:
            continue

        combined = " ".join(
            parts
        ).lower()

        # Skip header
        if (
            "plant" in combined
            and
            "quantity" in combined
        ):
            continue

        # Skip separator
        if all(
            re.fullmatch(
                r"-+",
                p.replace(
                    " ",
                    ""
                )
            )
            for p in parts
        ):
            continue

        plant_name = (
            parts[0]
            .replace("**", "")
            .strip()
        )

        matched_name = None

        for valid_name in valid_plants:

            if (
                plant_name.lower()
                ==
                valid_name.lower()
            ):

                matched_name = valid_name

                break

        if not matched_name:
            continue

        quantity_match = re.search(
            r"[\d,.]+",
            parts[1]
        )

        if not quantity_match:
            continue

        try:

            quantity = float(
                quantity_match
                .group()
                .replace(
                    ",",
                    ""
                )
            )

        except Exception:

            continue

        # Get official rate from CSV
        db_row = plants_df[
            plants_df["plant"]
            .astype(str)
            .str.strip()
            ==
            matched_name
        ]

        if db_row.empty:
            continue

        rate = float(
            db_row.iloc[0]["price"]
        )

        amount = (
            quantity * rate
        )

        rows.append(
            {
                "Plant": matched_name,
                "Quantity": quantity,
                "Unit": parts[2],
                "Rate (₹)": rate,
                "Amount (₹)": amount
            }
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Plant",
            "Quantity",
            "Unit",
            "Rate (₹)",
            "Amount (₹)"
        ]
    )


# ============================================================
# EXTRACT ONLY SELECTED QUOTATION ITEMS
# ============================================================

def extract_budget(
    landscape_result,
    selected_options,
    plant_df
):

    rows = []
    budget_started = False

    # These are the only quotation categories that can appear.
    # Everything else is deliberately excluded.
    allowed_categories = set()

    plant_options = {
        "Lawn",
        "Trees",
        "Boundary / Hedge",
        "Flowering Plants",
        "Ornamental Plants",
        "Palm / Tropical Theme",
        "Vastu Plants"
    }

    if any(option in plant_options for option in selected_options):
        if not plant_df.empty:
            allowed_categories.add("plant material & lawn grass")

    option_to_category = {
        "Soil Improvement": "soil improvement",
        "Irrigation System": "irrigation system",
        "Pathway": "pathway & border construction",
        "Pathway & Border Construction": "pathway & border construction",
        "Rock Garden": "rock garden",
        "Water Feature": "water feature",
        "Seating Area": "seating area",
        "Parking Landscape": "parking landscape",
        "Landscape Lighting": "landscape lighting",
        "Labour & Garden Execution": "labour & garden execution",
        "Contingency / Miscellaneous": "contingency / miscellaneous"
    }

    for option in selected_options:
        if option in option_to_category:
            allowed_categories.add(option_to_category[option])

    if not allowed_categories:
        return pd.DataFrame(
            columns=["Category", "Estimated Cost (₹)"]
        )

    for line in landscape_result.splitlines():

        stripped = line.strip()
        lower = stripped.lower()

        if "preliminary budget" in lower:
            budget_started = True
            continue

        if not budget_started:
            continue

        if (
            stripped.startswith("# ")
            and
            (
                "site verification" in lower
                or "customer-friendly" in lower
            )
        ):
            break

        if not stripped.startswith("|"):
            continue

        parts = [
            part.strip()
            for part in stripped.strip("|").split("|")
        ]

        if len(parts) < 2:
            continue

        combined = " ".join(parts).lower()

        if "category" in combined and "cost" in combined:
            continue

        if all(
            re.fullmatch(r"-+", p.replace(" ", ""))
            for p in parts
        ):
            continue

        category = (
            parts[0]
            .replace("**", "")
            .strip()
        )

        # Never import the AI's total row into the quotation items.
        if category.lower() == "total":
            continue

        normalized = category.lower()

        if normalized not in allowed_categories:
            continue

        amount_match = re.search(r"[\d,.]+", parts[1])

        if not amount_match:
            continue

        try:
            amount = float(
                amount_match.group().replace(",", "")
            )
        except Exception:
            continue

        rows.append(
            {
                "Category": category,
                "Estimated Cost (₹)": amount
            }
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Category",
            "Estimated Cost (₹)"
        ]
    )


# ============================================================
# PREPARE IMAGE FOR CLOUDFLARE
# ============================================================

def prepare_reference_image(
    uploaded_image
):

    image = Image.open(
        uploaded_image
    ).convert("RGB")

    # Cloudflare requires reference images
    # to be smaller than 512x512.

    image.thumbnail(
        (511, 511),
        Image.Resampling.LANCZOS
    )

    output = BytesIO()

    image.save(
        output,
        format="PNG"
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# CREATE VISUAL PROMPT
# ============================================================

def create_visual_prompt(
    project_name,
    plant_df
):

    # The image model receives ONLY the Plant BOQ.
    # It does not receive the complete landscape report.

    plant_lines = []

    if plant_df is not None and not plant_df.empty:

        for _, row in plant_df.iterrows():

            plant_name = str(row["Plant"]).strip()
            quantity = row["Quantity"]
            unit = str(row["Unit"]).strip()

            plant_lines.append(
                f"- {plant_name}: {quantity} {unit}"
            )

    if not plant_lines:

        raise Exception(
            "Plant BOQ is empty. Generate the landscape plan "
            "with at least one plant before creating the visualization."
        )

    plant_boq = "\n".join(plant_lines)

    return f"""
STRICT EXISTING-PHOTO PLANT-ONLY EDIT.

PROJECT:
{project_name}

============================================================
ONLY ALLOWED ADDITIONS
============================================================

The ONLY new visual elements allowed in the final image
are the plants listed in this Plant BOQ:

{plant_boq}

No other new object is allowed.

============================================================
CRITICAL INSTRUCTION
============================================================

Use the uploaded image as the exact base photograph.

KEEP THE ORIGINAL PHOTOGRAPH AS UNCHANGED AS POSSIBLE.

Do NOT create a new scene.
Do NOT redesign the property.
Do NOT reinterpret the property.
Do NOT beautify the property.
Do NOT generate an architectural visualization.

This is an IN-PLACE PHOTO EDIT.

Original photograph + ONLY the BOQ plants.

============================================================
ABSOLUTELY FORBIDDEN ADDITIONS
============================================================

DO NOT add:

- buildings
- houses
- rooms
- extensions
- structures
- pergolas
- gazebos
- sheds
- walls
- fences
- gates
- pathways
- roads
- paving
- tiles
- driveways
- seating
- benches
- tables
- chairs
- water features
- fountains
- ponds
- pools
- rock gardens
- decorative rocks
- lighting
- landscape lights
- statues
- pots
- planters
- furniture
- vehicles
- people
- animals
- signs
- decorations
- construction elements
- hardscape
- architectural changes
- any object that is not one of the BOQ plants

============================================================
PLANT SPECIES RULE
============================================================

Use ONLY the exact plant names listed in the BOQ.

Do not substitute similar species.

Do not invent plant species.

Do not add extra plant species.

If a plant is not in the BOQ, it MUST NOT appear as a
newly added plant.

============================================================
QUANTITY RULE
============================================================

Respect the BOQ quantities as closely as possible.

Do not create excessive vegetation.

Do not create a forest.

Do not duplicate plants unnecessarily.

For individual plants, approximately match the specified
quantity.

For area-based items such as lawn grass, add only the
corresponding amount of coverage to suitable existing
open ground.

============================================================
WHERE PLANTS MAY BE ADDED
============================================================

Plants may ONLY be added to clearly visible:

- bare soil
- open ground
- existing planting beds
- existing unlandscaped ground

Plants must be physically rooted in the existing ground.

Do not place plants:

- inside buildings
- on roofs
- on walls
- on roads
- floating in air
- on vehicles
- on existing structures

============================================================
EXISTING PHOTO MUST NOT CHANGE
============================================================

Preserve the existing:

- building
- walls
- doors
- windows
- roof
- boundary
- driveway
- road
- existing trees
- existing plants
- terrain
- construction materials
- objects
- sky
- clouds
- lighting
- shadows
- colors
- textures
- camera viewpoint
- perspective
- composition
- framing

Do not remove existing objects.

Do not move existing objects.

Do not replace existing objects.

Do not repaint existing objects.

Do not alter architecture.

Do not crop.

Do not zoom.

Do not change the camera angle.

============================================================
REALISM
============================================================

Only the newly added BOQ plants should be generated.

Match those plants to the original photograph's:

- perspective
- scale
- lighting
- sunlight direction
- shadows
- color temperature
- image quality

The result must look like the SAME photograph after only
the listed plants were planted.

============================================================
FINAL VALIDATION
============================================================

Before returning the image, verify:

1. Every new addition is a plant.
2. Every new plant is listed in the BOQ.
3. No building was added.
4. No structure was added.
5. No pathway was added.
6. No hardscape was added.
7. No furniture was added.
8. No decoration was added.
9. No unlisted plant species were added.
10. The original photograph remains visually unchanged
    except for the BOQ plant additions.

If any forbidden object appears, remove it.

FINAL OUTPUT:

ORIGINAL PHOTO
+
ONLY BOQ PLANTS

NOTHING ELSE.

No text.
No labels.
No logos.
No watermark.
"""



# ============================================================
# CLOUDFLARE FLUX IMAGE GENERATION
# ============================================================

def generate_cloudflare_image(
    uploaded_image,
    project_name,
    plant_df
):

    if not CLOUDFLARE_API_TOKEN:

        raise Exception(
            "CLOUDFLARE_API_TOKEN is missing "
            "from your .env file."
        )

    if not CLOUDFLARE_ACCOUNT_ID:

        raise Exception(
            "CLOUDFLARE_ACCOUNT_ID is missing "
            "from your .env file."
        )

    # ========================================================
    # ENDPOINT
    # ========================================================

    url = (
        "https://api.cloudflare.com/client/v4/"
        f"accounts/{CLOUDFLARE_ACCOUNT_ID}"
        f"/ai/run/{CLOUDFLARE_MODEL}"
    )

    # ========================================================
    # IMAGE
    # ========================================================

    image_bytes = (
        prepare_reference_image(
            uploaded_image
        )
    )

    # ========================================================
    # PROMPT
    # ========================================================

    prompt = create_visual_prompt(
        project_name,
        plant_df
    )

    # ========================================================
    # MULTIPART FORM DATA
    # ========================================================

    files = {

        "input_image_0": (
            "site.png",
            image_bytes,
            "image/png"
        )
    }

    data = {

        "prompt": prompt,

        "width": "1024",

        "height": "768",

        "guidance": "2.0"
    }

    headers = {

        "Authorization":
            f"Bearer {CLOUDFLARE_API_TOKEN}"
    }

    # ========================================================
    # REQUEST
    # ========================================================

    try:

        response = requests.post(
            url,
            headers=headers,
            data=data,
            files=files,
            timeout=180
        )

    except requests.exceptions.Timeout:

        raise Exception(
            "Cloudflare image generation timed out."
        )

    except requests.exceptions.RequestException as e:

        raise Exception(
            f"Cloudflare connection error:\n{e}"
        )

    # ========================================================
    # SUCCESS
    # ========================================================

    if response.status_code == 200:

        content_type = (
            response.headers
            .get(
                "content-type",
                ""
            )
            .lower()
        )

        # Direct image response
        if (
            "image/" in content_type
            or
            response.content[:4]
            ==
            b"\x89PNG"
            or
            response.content[:2]
            ==
            b"\xff\xd8"
        ):

            return response.content

        # JSON response
        try:

            response_json = (
                response.json()
            )

            result = (
                response_json.get(
                    "result",
                    {}
                )
            )

            image_base64 = None

            if isinstance(
                result,
                dict
            ):

                image_base64 = (
                    result.get(
                        "image"
                    )
                )

            if image_base64:

                # Remove data URI prefix
                if "," in image_base64:

                    image_base64 = (
                        image_base64
                        .split(
                            ",",
                            1
                        )[1]
                    )

                return base64.b64decode(
                    image_base64
                )

            raise Exception(
                "Cloudflare returned success "
                "but no image was found."
            )

        except Exception as e:

            raise Exception(
                "Unexpected Cloudflare response:\n"
                f"{e}"
            )

    # ========================================================
    # AUTHENTICATION
    # ========================================================

    if response.status_code == 401:

        try:
            details = response.json()
        except Exception:
            details = response.text

        raise Exception(
            "Cloudflare authentication failed.\n\n"
            f"{details}"
        )

    # ========================================================
    # MODEL ACCESS
    # ========================================================

    if response.status_code == 403:

        try:
            details = response.json()
        except Exception:
            details = response.text

        raise Exception(
            "Cloudflare denied access to "
            "FLUX.2 Klein 9B.\n\n"
            f"{details}"
        )

    # ========================================================
    # RATE LIMIT
    # ========================================================

    if response.status_code == 429:

        raise Exception(
            "Cloudflare image generation is temporarily unavailable "
            "because the account has reached its current AI limit. "
            "Please wait for the allowance to reset and try again."
        )

    # ========================================================
    # TEMPORARY CLOUDFLARE CAPACITY
    # ========================================================

    try:
        error_payload = response.json()
        error_code = str(
            error_payload.get("errors", [{}])[0].get("code", "")
        )
        error_message = str(
            error_payload.get("errors", [{}])[0].get("message", "")
        ).lower()
    except Exception:
        error_code = ""
        error_message = response.text.lower()

    if (
        error_code == "3040"
        or "capacity temporarily exceeded" in error_message
        or "capacity" in error_message
        and response.status_code >= 500
    ):
        raise Exception(
            "Cloudflare image generation is temporarily at capacity. "
            "Your Greenscape project and Plant BOQ are fine. "
            "Please wait a little and click Generate Site Visualisation again."
        )

    # ========================================================
    # OTHER CLOUDFLARE ERRORS
    # ========================================================

    try:
        details = response.json()
    except Exception:
        details = response.text

    raise Exception(
        f"Cloudflare image generation failed "
        f"(HTTP {response.status_code}).\n\n"
        f"{details}"
    )


# ============================================================
# EXCEL FILE
# ============================================================

def create_excel_file(
    project_name,
    location,
    area,
    landscape_result,
    plant_df,
    budget_df,
    selected_options=None
):

    selected_options = selected_options or []

    workbook = Workbook()

    # --------------------------------------------------------
    # PROFESSIONAL THEME
    # --------------------------------------------------------

    title_fill = PatternFill(
        "solid",
        fgColor="1F4E3D"
    )

    section_fill = PatternFill(
        "solid",
        fgColor="D9EAD3"
    )

    header_fill = PatternFill(
        "solid",
        fgColor="2F6B4F"
    )

    total_fill = PatternFill(
        "solid",
        fgColor="E2F0D9"
    )

    white_font = Font(
        color="FFFFFF",
        bold=True,
        size=12
    )

    title_font = Font(
        color="FFFFFF",
        bold=True,
        size=18
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    bold_font = Font(
        bold=True
    )

    thin_gray = Side(
        style="thin",
        color="D9E1E8"
    )

    border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray
    )

    currency_format = '₹#,##0.00'
    number_format = '#,##0.00'

    def style_title(ws, cell_range, text):
        ws.merge_cells(cell_range)
        cell = ws[cell_range.split(":")[0]]
        cell.value = text
        cell.fill = title_fill
        cell.font = title_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        ws.row_dimensions[cell.row].height = 30

    def style_header_row(ws, row, start_col, end_col):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            cell.border = border

    def apply_borders(ws):
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                cell.border = border
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )

    def auto_width(ws, max_width=55):
        for column_cells in ws.columns:
            if not column_cells:
                continue

            column_letter = get_column_letter(
                column_cells[0].column
            )

            max_length = 0

            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            ws.column_dimensions[column_letter].width = min(
                max(max_length + 3, 12),
                max_width
            )

    # --------------------------------------------------------
    # 1. PROJECT SUMMARY
    # --------------------------------------------------------

    ws1 = workbook.active
    ws1.title = "Quotation Summary"

    style_title(
        ws1,
        "A1:F1",
        "GREENSCAPE | LANDSCAPE PLANNING & QUOTATION"
    )

    summary_rows = [
        ("Project Name", project_name),
        ("Location", location),
        ("Area (sq.ft)", area),
    ]

    row = 3

    for label, value in summary_rows:
        ws1.cell(row=row, column=1, value=label)
        ws1.cell(row=row, column=1).font = bold_font
        ws1.cell(row=row, column=2, value=value)
        ws1.merge_cells(
            start_row=row,
            start_column=2,
            end_row=row,
            end_column=6
        )
        row += 1

    ws1.cell(row=7, column=1, value="SELECTED PROJECT SCOPE")
    ws1.cell(row=7, column=1).font = bold_font
    ws1.cell(row=7, column=1).fill = section_fill

    ws1.merge_cells("A7:F7")

    scope_text = (
        " • ".join(selected_options)
        if selected_options
        else "No optional scope selected"
    )

    ws1.cell(row=8, column=1, value=scope_text)
    ws1.merge_cells("A8:F9")
    ws1["A8"].alignment = Alignment(
        vertical="top",
        wrap_text=True
    )

    plant_total = (
        float(plant_df["Amount (₹)"].sum())
        if not plant_df.empty
        else 0
    )

    budget_total = (
        float(budget_df["Estimated Cost (₹)"].sum())
        if not budget_df.empty
        else 0
    )

    ws1.cell(row=11, column=1, value="Plant BOQ Total")
    ws1.cell(row=11, column=2, value=plant_total)
    ws1.cell(row=12, column=1, value="Selected Work Total")
    ws1.cell(row=12, column=2, value=budget_total)
    ws1.cell(row=13, column=1, value="Estimated Project Total")
    ws1.cell(row=13, column=2, value=plant_total + budget_total)

    for r in range(11, 14):
        ws1.cell(row=r, column=1).font = bold_font
        ws1.cell(row=r, column=2).number_format = currency_format

    ws1.cell(row=13, column=1).fill = total_fill
    ws1.cell(row=13, column=2).fill = total_fill
    ws1.cell(row=13, column=1).font = Font(bold=True, size=12)
    ws1.cell(row=13, column=2).font = Font(bold=True, size=12)

    ws1.freeze_panes = "A3"
    ws1.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # 2. SELECTED SCOPE
    # --------------------------------------------------------

    ws_scope = workbook.create_sheet("Selected Scope")

    style_title(
        ws_scope,
        "A1:C1",
        "SELECTED LANDSCAPE & HARDSCAPE SCOPE"
    )

    scope_headers = [
        "Sr. No.",
        "Scope Type",
        "Selected Item"
    ]

    for col, header in enumerate(scope_headers, start=1):
        ws_scope.cell(
            row=3,
            column=col,
            value=header
        )

    style_header_row(ws_scope, 3, 1, 3)

    landscape_scope = {
        "Lawn",
        "Trees",
        "Boundary / Hedge",
        "Flowering Plants",
        "Ornamental Plants",
        "Palm / Tropical Theme",
        "Vastu Plants",
        "Soil Improvement",
        "Irrigation System",
    }

    hardscape_scope = {
        "Pathway",
        "Rock Garden",
        "Water Feature",
        "Seating Area",
        "Parking Landscape",
        "Landscape Lighting",
        "Children's Area",
    }

    for idx, item in enumerate(
        selected_options,
        start=1
    ):
        if item in landscape_scope:
            scope_type = "Landscape"
        elif item in hardscape_scope:
            scope_type = "Hardscape"
        else:
            scope_type = "Quotation Add-on"

        ws_scope.cell(idx + 3, 1, idx)
        ws_scope.cell(idx + 3, 2, scope_type)
        ws_scope.cell(idx + 3, 3, item)

    ws_scope.freeze_panes = "A4"
    ws_scope.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # 3. PLANT BOQ
    # --------------------------------------------------------

    ws2 = workbook.create_sheet("Plant BOQ")

    style_title(
        ws2,
        "A1:E1",
        "PLANT MATERIAL BOQ"
    )

    headers = [
        "Plant",
        "Quantity",
        "Unit",
        "Rate (₹)",
        "Amount (₹)"
    ]

    for col, header in enumerate(headers, start=1):
        ws2.cell(
            row=3,
            column=col,
            value=header
        )

    style_header_row(ws2, 3, 1, 5)

    for row_index, (_, item) in enumerate(
        plant_df.iterrows(),
        start=4
    ):
        ws2.cell(row_index, 1, item["Plant"])
        ws2.cell(row_index, 2, item["Quantity"])
        ws2.cell(row_index, 3, item["Unit"])
        ws2.cell(row_index, 4, item["Rate (₹)"])
        ws2.cell(row_index, 5, item["Amount (₹)"])

        ws2.cell(row_index, 4).number_format = currency_format
        ws2.cell(row_index, 5).number_format = currency_format

    if not plant_df.empty:
        total_row = len(plant_df) + 4

        ws2.cell(total_row, 4, "TOTAL")
        ws2.cell(
            total_row,
            5,
            f"=SUM(E4:E{total_row - 1})"
        )

        for col in range(4, 6):
            ws2.cell(total_row, col).font = bold_font
            ws2.cell(total_row, col).fill = total_fill

        ws2.cell(
            total_row,
            5
        ).number_format = currency_format

    ws2.freeze_panes = "A4"
    ws2.auto_filter.ref = (
        f"A3:E{max(3, len(plant_df) + 3)}"
    )
    ws2.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # 4. QUOTATION
    # --------------------------------------------------------

    ws3 = workbook.create_sheet("Quotation")

    style_title(
        ws3,
        "A1:D1",
        "LANDSCAPE & HARDSCAPE QUOTATION"
    )

    quotation_headers = [
        "Sr. No.",
        "Category",
        "Description",
        "Estimated Cost (₹)"
    ]

    for col, header in enumerate(
        quotation_headers,
        start=1
    ):
        ws3.cell(
            row=3,
            column=col,
            value=header
        )

    style_header_row(ws3, 3, 1, 4)

    qrow = 4

    # Plant BOQ appears only when plant items were selected.
    if not plant_df.empty:
        ws3.cell(qrow, 1, 1)
        ws3.cell(qrow, 2, "Landscape")
        ws3.cell(
            qrow,
            3,
            "Plant Material & Lawn Grass"
        )
        ws3.cell(qrow, 4, plant_total)
        ws3.cell(qrow, 4).number_format = currency_format
        qrow += 1

    for idx, (_, item) in enumerate(
        budget_df.iterrows(),
        start=1
    ):
        ws3.cell(qrow, 1, qrow - 3)
        category = str(item["Category"])

        category_type = (
            "Hardscape"
            if any(
                x in category.lower()
                for x in [
                    "pathway",
                    "rock garden",
                    "water feature",
                    "seating",
                    "parking",
                    "lighting"
                ]
            )
            else "Landscape / Add-on"
        )

        ws3.cell(qrow, 2, category_type)
        ws3.cell(qrow, 3, category)
        ws3.cell(
            qrow,
            4,
            item["Estimated Cost (₹)"]
        )
        ws3.cell(qrow, 4).number_format = currency_format
        qrow += 1

    total_row = qrow + 1

    ws3.cell(total_row, 3, "GRAND TOTAL")
    ws3.cell(
        total_row,
        4,
        f"=SUM(D4:D{qrow - 1})"
    )
    ws3.cell(total_row, 3).font = bold_font
    ws3.cell(total_row, 4).font = bold_font
    ws3.cell(total_row, 3).fill = total_fill
    ws3.cell(total_row, 4).fill = total_fill
    ws3.cell(total_row, 4).number_format = currency_format

    ws3.freeze_panes = "A4"
    ws3.auto_filter.ref = f"A3:D{max(3, qrow - 1)}"
    ws3.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # 5. LANDSCAPE REPORT
    # --------------------------------------------------------

    ws4 = workbook.create_sheet("Landscape Report")

    style_title(
        ws4,
        "A1:F1",
        "GREENSCAPE | LANDSCAPE PLANNING REPORT"
    )

    row = 3

    for line in landscape_result.splitlines():
        if line.strip():
            ws4.cell(row, 1, line.strip())
            ws4.merge_cells(
                start_row=row,
                start_column=1,
                end_row=row,
                end_column=6
            )
            row += 1

    ws4.column_dimensions["A"].width = 28
    for col in "BCDEF":
        ws4.column_dimensions[col].width = 20

    ws4.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # 6. NOTES
    # --------------------------------------------------------

    ws5 = workbook.create_sheet("Notes")

    style_title(
        ws5,
        "A1:F1",
        "IMPORTANT NOTES"
    )

    notes = [
        "This is a preliminary planning estimate prepared through Greenscape.",
        "Final quantities must be verified at site.",
        "Final rates may change according to plant size, quality, transportation and execution.",
        "Site measurements must be confirmed before final quotation.",
        "Plant recommendations are restricted to the supplied plant database.",
        "Only the selected landscape and hardscape scope is included in the quotation.",
        "Final execution should be reviewed by a landscape professional before work begins."
    ]

    for index, note in enumerate(notes, start=3):
        ws5.cell(index, 1, f"{index - 2}.")
        ws5.cell(index, 2, note)
        ws5.merge_cells(
            start_row=index,
            start_column=2,
            end_row=index,
            end_column=6
        )

    ws5.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # GLOBAL FORMATTING
    # --------------------------------------------------------

    for ws in workbook.worksheets:
        apply_borders(ws)
        auto_width(ws)

        # Keep the main report and notes readable.
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )

    # Restore title styling after global formatting.
    for ws in workbook.worksheets:
        for merged_range in ws.merged_cells.ranges:
            first_cell = ws.cell(
                merged_range.min_row,
                merged_range.min_col
            )
            if first_cell.row == 1:
                first_cell.fill = title_fill
                first_cell.font = title_font
                first_cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.freeze_panes = ws.freeze_panes

    # Make key quotation columns wide enough.
    for ws in [ws3]:
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 42
        ws.column_dimensions["D"].width = 20

    for ws in [ws2]:
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 16
        ws.column_dimensions["E"].width = 18

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output


# ============================================================
# PROJECT DETAILS
# ============================================================

st.markdown(
    '<div class="gs-section">'
    'Project Details'
    '</div>',
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

with col1:

    project_name = st.text_input(
        "Project Name",
        placeholder="Example: Villa Garden Project"
    )

with col2:

    location = st.text_input(
        "Location",
        placeholder="Example: Pune, Maharashtra"
    )

with col3:

    area = st.number_input(
        "Area (sq.ft)",
        min_value=1.0,
        value=1000.0,
        step=50.0
    )


# ============================================================
# SITE CONDITIONS
# ============================================================

st.markdown(
    '<div class="gs-section">'
    'Site Conditions'
    '</div>',
    unsafe_allow_html=True
)

col1, col2, col3, col4 = st.columns(4)

with col1:

    sunlight = st.selectbox(
        "Sunlight",
        [
            "Full Sun",
            "Partial Shade",
            "Mostly Shade",
            "Mixed"
        ]
    )

with col2:

    soil = st.selectbox(
        "Soil Type",
        [
            "Loamy",
            "Clay",
            "Sandy",
            "Rocky",
            "Red Soil",
            "Unknown"
        ]
    )

with col3:

    water = st.selectbox(
        "Water Availability",
        [
            "High",
            "Medium",
            "Low",
            "Unknown"
        ]
    )

with col4:

    maintenance = st.selectbox(
        "Maintenance Level",
        [
            "Very Low",
            "Low",
            "Medium",
            "High"
        ]
    )


# ============================================================
# LANDSCAPE REQUIREMENTS / QUOTATION SCOPE
# ============================================================

st.markdown(
    '<div class="gs-section">'
    'Scope & Requirements'
    '</div>',
    unsafe_allow_html=True
)

st.info(
    "Select only the work required for this project. The selected scope controls the planning proposal, estimate and quotation."
)

# ------------------------------------------------------------
# SEPARATE LANDSCAPE / HARDSCAPE / ADD-ON SCOPE
# ------------------------------------------------------------

landscape_options = [
    "Lawn",
    "Trees",
    "Boundary / Hedge",
    "Flowering Plants",
    "Ornamental Plants",
    "Palm / Tropical Theme",
    "Vastu Plants",
    "Soil Improvement",
    "Irrigation System",
]

hardscape_options = [
    "Pathway",
    "Rock Garden",
    "Water Feature",
    "Seating Area",
    "Parking Landscape",
    "Landscape Lighting",
    "Children's Area",
]

quotation_addon_options = [
    "Labour & Garden Execution",
    "Contingency / Miscellaneous",
]

st.markdown('<div class="gs-subsection">Landscape</div>', unsafe_allow_html=True)

selected_landscape = st.multiselect(
    "Select Landscape Items",
    options=landscape_options,
    default=[],
    placeholder="Choose landscape items..."
)

if selected_landscape:
    st.caption("Selected Landscape: " + ", ".join(selected_landscape))
else:
    st.caption("No landscape items selected.")

st.markdown('<div class="gs-subsection">Hardscape</div>', unsafe_allow_html=True)

selected_hardscape = st.multiselect(
    "Select Hardscape Items",
    options=hardscape_options,
    default=[],
    placeholder="Choose hardscape items..."
)

if selected_hardscape:
    st.caption("Selected Hardscape: " + ", ".join(selected_hardscape))
else:
    st.caption("No hardscape items selected.")

st.markdown('<div class="gs-subsection">Quotation Add-ons</div>', unsafe_allow_html=True)

selected_addons = st.multiselect(
    "Select Additional Quotation Items",
    options=quotation_addon_options,
    default=[],
    placeholder="Choose quotation add-ons..."
)

if selected_addons:
    st.caption("Selected Add-ons: " + ", ".join(selected_addons))
else:
    st.caption("No quotation add-ons selected.")

selected_options = (
    selected_landscape
    + selected_hardscape
    + selected_addons
)

if selected_options:
    st.success(
        "Selected scope: " + " • ".join(selected_options)
    )
else:
    st.warning(
        "No optional scope selected. The quotation will contain no optional work items."
    )

requirements = selected_options.copy()

if requirements:
    requirements_text = "\n".join(
        f"- {item}"
        for item in requirements
    )
else:
    requirements_text = "- No optional landscape or hardscape work selected"


# ============================================================
# SELECTED SCOPE SUMMARY
# ============================================================

with st.expander("View Selected Scope", expanded=False):

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("**🌱 Landscape**")
        if selected_landscape:
            for item in selected_landscape:
                st.write("•", item)
        else:
            st.caption("None selected")

    with c2:
        st.markdown("**🧱 Hardscape**")
        if selected_hardscape:
            for item in selected_hardscape:
                st.write("•", item)
        else:
            st.caption("None selected")

    with c3:
        st.markdown("**💼 Add-ons**")
        if selected_addons:
            for item in selected_addons:
                st.write("•", item)
        else:
            st.caption("None selected")


# ============================================================
# ADDITIONAL REQUIREMENTS
# ============================================================

additional_requirements = st.text_area(
    "Additional Requirements",
    placeholder=(
        "Example: Keep the center open, "
        "add flowering plants near entrance, "
        "use low-maintenance plants..."
    ),
    height=120
)


# ============================================================
# IMAGE UPLOAD
# ============================================================

st.markdown(
    '<div class="gs-section">'
    'Site Photograph'
    '</div>',
    unsafe_allow_html=True
)

uploaded_image = st.file_uploader(
    "Upload the existing site photograph",
    type=[
        "jpg",
        "jpeg",
        "png",
        "webp"
    ]
)

if uploaded_image:

    try:

        preview = Image.open(
            uploaded_image
        )

        st.image(
            preview,
            caption="Existing site",
            width="stretch"
        )

    except Exception as e:

        st.error(
            f"Could not preview image: {e}"
        )


# ============================================================
# GENERATE LANDSCAPE PLAN
# ============================================================

st.markdown(
    '<div class="gs-section">'
    'Planning Proposal'
    '</div>',
    unsafe_allow_html=True
)

generate_plan = st.button(
    "Prepare Planning Proposal",
    type="primary",
    width="stretch"
)


if generate_plan:

    if not project_name.strip():

        st.warning(
            "Please enter a project name."
        )

        st.stop()

    if not location.strip():

        st.warning(
            "Please enter the project location."
        )

        st.stop()

    if not uploaded_image:

        st.warning(
            "Please upload a site image."
        )

        st.stop()

    try:

        with st.spinner(
            "Analysing the site and preparing the planning proposal..."
        ):

            result = generate_landscape_plan(
                project_name,
                location,
                area,
                sunlight,
                soil,
                water,
                maintenance,
                requirements_text,
                additional_requirements,
                uploaded_image
            )

        st.session_state.landscape_result = (
            result
        )

        st.session_state.generated_image = (
            None
        )

        st.success(
            "Planning proposal prepared successfully."
        )

    except Exception as e:

        st.error(
            str(e)
        )


# ============================================================
# DISPLAY LANDSCAPE RESULT
# ============================================================

if st.session_state.landscape_result:

    landscape_result = (
        st.session_state.landscape_result
    )

    st.markdown(
        "---"
    )

    st.markdown(
        '<div class="gs-section">'
        'Planning Report'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        landscape_result
    )


    # ========================================================
    # EXTRACT BOQ
    # ========================================================

    plant_df = extract_plant_boq(
        landscape_result
    )

    budget_df = extract_budget(
        landscape_result,
        selected_options,
        plant_df
    )

    st.session_state.plant_df = (
        plant_df
    )

    st.session_state.budget_df = (
        budget_df
    )


    # ========================================================
    # PLANT BOQ
    # ========================================================

    st.markdown(
        '<div class="gs-section">'
        'Plant Schedule'
        '</div>',
        unsafe_allow_html=True
    )

    if not plant_df.empty:

        display_plant_df = (
            plant_df.copy()
        )

        display_plant_df[
            "Quantity"
        ] = (
            display_plant_df[
                "Quantity"
            ].round(2)
        )

        display_plant_df[
            "Rate (₹)"
        ] = (
            display_plant_df[
                "Rate (₹)"
            ].round(2)
        )

        display_plant_df[
            "Amount (₹)"
        ] = (
            display_plant_df[
                "Amount (₹)"
            ].round(2)
        )

        st.dataframe(
            display_plant_df,
            width="stretch",
            hide_index=True
        )

        plant_total = (
            plant_df[
                "Amount (₹)"
            ].sum()
        )

        st.metric(
            "Plant Material & Lawn Cost",
            f"₹{plant_total:,.2f}"
        )

    else:

        st.warning(
            "No plant BOQ could be extracted."
        )


    # ========================================================
    # BUDGET
    # ========================================================

    st.markdown(
        '<div class="gs-section">'
        'Preliminary Estimate'
        '</div>',
        unsafe_allow_html=True
    )

    if not budget_df.empty:

        st.dataframe(
            budget_df,
            width="stretch",
            hide_index=True
        )

        total_budget = (
            budget_df[
                "Estimated Cost (₹)"
            ].sum()
        )

        st.metric(
            "Estimated Project Cost",
            f"₹{total_budget:,.2f}"
        )

    else:

        st.info(
            "No quotation items are selected, so no additional budget items are shown."
        )


    # ========================================================
    # EXCEL
    # ========================================================

    st.markdown(
        '<div class="gs-section">'
        'BOQ & Quotation'
        '</div>',
        unsafe_allow_html=True
    )

    excel_file = create_excel_file(
        project_name,
        location,
        area,
        landscape_result,
        plant_df,
        budget_df,
        selected_options
    )

    safe_name = (
        project_name
        .strip()
        .replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
    )

    st.download_button(
        label="Download BOQ & Quotation",
        data=excel_file,
        file_name=(
            f"{safe_name}_"
            "Landscape_Quotation.xlsx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        width="stretch"
    )


    # ========================================================
    # AI VISUALIZATION
    # ========================================================

    st.markdown(
        "---"
    )

    st.markdown(
        '<div class="gs-section">'
        'Site Visualisation'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Create an in-place site visualisation using ONLY the plants "
        "listed in the Plant Schedule. No buildings, structures, "
        "hardscape, furniture, or other additions are allowed."
    )

    generate_visual = st.button(
        "Generate Site Visualisation",
        width="stretch"
    )

    st.caption(
        "Visualisation uses the uploaded photograph and Plant BOQ only. "
        "If Cloudflare is temporarily at capacity, try again later."
    )

    if generate_visual:

        if plant_df.empty:

            st.error(
                "Cannot generate visualization because the Plant BOQ is empty."
            )

        elif not CLOUDFLARE_API_TOKEN:

            st.error(
                "CLOUDFLARE_API_TOKEN is missing "
                "from your .env file."
            )

        elif not CLOUDFLARE_ACCOUNT_ID:

            st.error(
                "CLOUDFLARE_ACCOUNT_ID is missing "
                "from your .env file."
            )

        else:

            try:

                with st.spinner(
                    "Preparing the site visualisation..."
                ):

                    generated_image = (
                        generate_cloudflare_image(
                            uploaded_image,
                            project_name,
                            plant_df
                        )
                    )

                st.session_state.generated_image = (
                    generated_image
                )

                st.success(
                    "Site visualisation prepared successfully."
                )

            except Exception as e:

                st.error(
                    f"Site visualisation could not be generated.\n\n{e}"
                )


    # ========================================================
    # DISPLAY GENERATED IMAGE
    # ========================================================

    if st.session_state.generated_image:

        st.markdown(
            "### Proposed Site View"
        )

        generated_image = (
            st.session_state.generated_image
        )

        st.image(
            generated_image,
            caption="Proposed site view — selected plants only",
            width="stretch"
        )

        st.download_button(
            label="Download Visualisation",
            data=generated_image,
            file_name=(
                f"{safe_name}_"
                "AI_Landscape.png"
            ),
            mime="image/png",
            width="stretch"
        )


# ============================================================
# ============================================================
# GREENSCAPE FOOTER
# ============================================================

st.markdown(
    """
    <div class="gs-footer">
        Greenscape · Landscape Planning & Estimation
        <br>
        Project planning interface
    </div>
    """,
    unsafe_allow_html=True
)
