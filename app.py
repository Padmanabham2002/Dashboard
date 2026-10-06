import os
import re
import hashlib
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Competitive Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# CONSTANTS
# ============================================================

import os
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(BASE_DIR, "output.xlsx")

STATUS_RELEVANT = "Relevant information found"
STATUS_NOT_RELEVANT = "No relevant information found"
STATUS_NOT_PROVIDED = "Not Provided"


# ============================================================
# SESSION STATE
# ============================================================

if "page" not in st.session_state:
    st.session_state.page = "Executive Overview"


# ============================================================
# CSS ONLY
# ============================================================

st.markdown(
    """
    <style>

    /* ========================================================
       GLOBAL APPLICATION BACKGROUND
       ======================================================== */

    html,
    body,
    [data-testid="stApp"],
    [data-testid="stAppViewContainer"] {

        background:
            linear-gradient(
                135deg,
                #ffffff 0%,
                #fbfdff 28%,
                #fcfbff 55%,
                #fffdf9 100%
            ) !important;

        overflow-x: hidden !important;
    }


    /* ========================================================
       HIDE STREAMLIT HEADER / TOOLBAR
       ======================================================== */

    [data-testid="stHeader"] {

        background: transparent !important;

        height: 0 !important;

        min-height: 0 !important;
    }


    [data-testid="stToolbar"] {

        display: none !important;

        visibility: hidden !important;
    }


    [data-testid="stDecoration"] {

        display: none !important;
    }


    [data-testid="stAppDeployButton"] {

        display: none !important;

        visibility: hidden !important;
    }


    #MainMenu {

        visibility: hidden !important;
    }


    footer {

        visibility: hidden !important;
    }


    /* ========================================================
       BEAUTIFUL LIGHT BACKGROUND - CYAN GLOW
       ======================================================== */

    [data-testid="stApp"]::before {

        content: "";

        position: fixed;

        width: 520px;

        height: 520px;

        left: -210px;

        top: 30px;

        border-radius: 50%;

        background:
            radial-gradient(
                circle,
                rgba(61, 190, 232, 0.14) 0%,
                rgba(61, 190, 232, 0.07) 28%,
                rgba(61, 190, 232, 0.025) 52%,
                transparent 74%
            );

        filter: blur(8px);

        pointer-events: none !important;

        z-index: 0 !important;

        animation:
            cyanBackgroundMove
            16s ease-in-out infinite;
    }


    @keyframes cyanBackgroundMove {

        0% {

            transform:
                translate3d(0, 0, 0)
                scale(0.95);
        }

        25% {

            transform:
                translate3d(90px, 40px, 0)
                scale(1.08);
        }

        50% {

            transform:
                translate3d(150px, 130px, 0)
                scale(0.92);
        }

        75% {

            transform:
                translate3d(60px, 190px, 0)
                scale(1.08);
        }

        100% {

            transform:
                translate3d(0, 0, 0)
                scale(0.95);
        }
    }


    /* ========================================================
       BEAUTIFUL LIGHT BACKGROUND - LAVENDER / PEACH
       ======================================================== */

    [data-testid="stApp"]::after {

        content: "";

        position: fixed;

        width: 600px;

        height: 600px;

        right: -250px;

        bottom: -160px;

        border-radius: 50%;

        background:
            radial-gradient(
                circle,
                rgba(190, 145, 245, 0.10) 0%,
                rgba(255, 181, 105, 0.08) 28%,
                rgba(190, 145, 245, 0.035) 52%,
                transparent 75%
            );

        filter: blur(10px);

        pointer-events: none !important;

        z-index: 0 !important;

        animation:
            purpleBackgroundMove
            19s ease-in-out infinite;
    }


    @keyframes purpleBackgroundMove {

        0% {

            transform:
                translate3d(0, 0, 0)
                scale(1);
        }

        25% {

            transform:
                translate3d(-80px, -50px, 0)
                scale(1.10);
        }

        50% {

            transform:
                translate3d(-150px, -130px, 0)
                scale(0.92);
        }

        75% {

            transform:
                translate3d(-50px, -80px, 0)
                scale(1.08);
        }

        100% {

            transform:
                translate3d(0, 0, 0)
                scale(1);
        }
    }


    /* ========================================================
       THIRD SMALL BACKGROUND LIGHT
       ======================================================== */

    .block-container::after {

        content: "";

        position: fixed;

        width: 180px;

        height: 180px;

        right: 14%;

        top: 115px;

        border-radius: 50%;

        background:
            radial-gradient(
                circle,
                rgba(96, 165, 250, 0.10) 0%,
                rgba(129, 140, 248, 0.045) 40%,
                transparent 72%
            );

        filter: blur(3px);

        pointer-events: none !important;

        z-index: 0 !important;

        animation:
            smallBackgroundMove
            9s ease-in-out infinite;
    }


    @keyframes smallBackgroundMove {

        0% {

            transform:
                translate3d(0, 0, 0)
                scale(1);
        }

        25% {

            transform:
                translate3d(-35px, 25px, 0)
                scale(1.12);
        }

        50% {

            transform:
                translate3d(-10px, 65px, 0)
                scale(0.90);
        }

        75% {

            transform:
                translate3d(40px, 20px, 0)
                scale(1.10);
        }

        100% {

            transform:
                translate3d(0, 0, 0)
                scale(1);
        }
    }


    /* ========================================================
       A4 DASHBOARD PAGE
       ======================================================== */

    .block-container {

        max-width: 900px !important;

        min-height: 1000px !important;

        margin-left: auto !important;

        margin-right: auto !important;

        padding-top: 0.45rem !important;

        padding-bottom: 3rem !important;

        padding-left: 1rem !important;

        padding-right: 1rem !important;

        position: relative !important;

        z-index: 2 !important;

        perspective: 1400px;

        background:
            rgba(255, 255, 255, 0.90) !important;

        border-left:
            1px solid
            rgba(220, 225, 232, 0.70);

        border-right:
            1px solid
            rgba(220, 225, 232, 0.70);

        box-shadow:
            0 0 35px
            rgba(30, 45, 70, 0.045);

        overflow: visible !important;

        backdrop-filter:
            blur(8px);
    }


    /* ========================================================
       SUBTLE MOVING GRID
       ======================================================== */

    .block-container {

        background-image:

            linear-gradient(
                rgba(79, 70, 229, 0.008) 1px,
                transparent 1px
            ),

            linear-gradient(
                90deg,
                rgba(79, 70, 229, 0.008) 1px,
                transparent 1px
            );

        background-size:
            48px 48px;

        animation:
            dashboardGridMove
            30s linear infinite;
    }


    @keyframes dashboardGridMove {

        0% {

            background-position:
                0 0,
                0 0;
        }

        100% {

            background-position:
                48px 48px,
                48px 48px;
        }
    }


    /* ========================================================
       SUVERA WATERMARK
       
       IMPORTANT:
       - fixed to browser viewport
       - diagonal
       - end-to-end
       - does NOT scroll
       - does NOT animate
       ======================================================== */

    [data-testid="stApp"] .main::before {

        content: "SUVERA";

        position: fixed !important;

        left: 50vw !important;

        top: 54vh !important;

        width: 150vw !important;

        height: 190px !important;

        display: flex !important;

        align-items: center !important;

        justify-content: center !important;

        transform:
            translate(-50%, -50%)
            rotate(-31deg) !important;

        transform-origin:
            center center !important;

        white-space: nowrap !important;

        pointer-events: none !important;

        user-select: none !important;

        z-index: 1 !important;

        font-size: 185px !important;

        line-height: 1 !important;

        font-weight: 900 !important;

        letter-spacing: 24px !important;

        color:
            rgba(83, 96, 120, 0.045) !important;

        text-shadow:
            0 2px 3px
            rgba(255, 255, 255, 0.55) !important;

        animation: none !important;
    }


    /* ========================================================
       KEEP REAL DASHBOARD CONTENT ABOVE WATERMARK
       ======================================================== */

    .block-container > * {

        position: relative;

        z-index: 3;
    }


    /* ========================================================
       HEADER
       ======================================================== */

    [data-testid="stHeading"] {

        text-align: center !important;

        margin-top: 0 !important;

        margin-bottom: 0 !important;

        position: relative;

        z-index: 10;

        perspective: 1000px;
    }


    /* ========================================================
       HEADING - LEFT COLORED 3D OBJECT
       ======================================================== */

    [data-testid="stHeading"]::before {

        content: "◆";

        position: absolute;

        left: 15%;

        top: 4px;

        font-size: 31px;

        font-weight: 900;

        color:
            #24c7e8;

        opacity:
            0.82;

        text-shadow:

            5px 5px 0
            rgba(65, 125, 255, 0.18),

            9px 9px 15px
            rgba(30, 185, 225, 0.25),

            0 0 20px
            rgba(30, 195, 235, 0.34);

        transform-style:
            preserve-3d;

        animation:
            headingObjectLeft
            5.5s ease-in-out infinite;
    }


    @keyframes headingObjectLeft {

        0% {

            transform:
                translate3d(0,0,0)
                rotateX(0deg)
                rotateY(0deg)
                rotateZ(0deg)
                scale(1);
        }

        20% {

            transform:
                translate3d(-10px,-10px,20px)
                rotateX(35deg)
                rotateY(45deg)
                rotateZ(15deg)
                scale(1.05);
        }

        40% {

            transform:
                translate3d(5px,8px,35px)
                rotateX(65deg)
                rotateY(90deg)
                rotateZ(25deg)
                scale(0.90);
        }

        60% {

            transform:
                translate3d(14px,-5px,20px)
                rotateX(90deg)
                rotateY(145deg)
                rotateZ(35deg)
                scale(1.08);
        }

        80% {

            transform:
                translate3d(-5px,8px,12px)
                rotateX(40deg)
                rotateY(210deg)
                rotateZ(15deg)
                scale(0.96);
        }

        100% {

            transform:
                translate3d(0,0,0)
                rotateX(0deg)
                rotateY(360deg)
                rotateZ(0deg)
                scale(1);
        }
    }


    /* ========================================================
       HEADING - RIGHT COLORED 3D OBJECT
       ======================================================== */

    [data-testid="stHeading"]::after {

        content: "✦";

        position: absolute;

        right: 16%;

        top: 3px;

        font-size: 32px;

        font-weight: 900;

        color:
            #ffad4a;

        opacity:
            0.84;

        text-shadow:

            -5px 5px 0
            rgba(160, 90, 240, 0.17),

            -9px 9px 15px
            rgba(255, 150, 70, 0.27),

            0 0 20px
            rgba(255, 175, 75, 0.34);

        transform-style:
            preserve-3d;

        animation:
            headingObjectRight
            6.5s ease-in-out infinite;
    }


    @keyframes headingObjectRight {

        0% {

            transform:
                translate3d(0,0,0)
                rotateX(0deg)
                rotateY(0deg)
                rotateZ(0deg)
                scale(1);
        }

        20% {

            transform:
                translate3d(12px,7px,20px)
                rotateX(-30deg)
                rotateY(-45deg)
                rotateZ(-12deg)
                scale(1.05);
        }

        40% {

            transform:
                translate3d(-5px,-12px,35px)
                rotateX(-65deg)
                rotateY(-90deg)
                rotateZ(-22deg)
                scale(0.90);
        }

        60% {

            transform:
                translate3d(-14px,5px,20px)
                rotateX(-90deg)
                rotateY(-145deg)
                rotateZ(-35deg)
                scale(1.08);
        }

        80% {

            transform:
                translate3d(6px,-6px,12px)
                rotateX(-40deg)
                rotateY(-210deg)
                rotateZ(-15deg)
                scale(0.96);
        }

        100% {

            transform:
                translate3d(0,0,0)
                rotateX(0deg)
                rotateY(-360deg)
                rotateZ(0deg)
                scale(1);
        }
    }


    /* ========================================================
       MAIN TITLE
       ======================================================== */

    [data-testid="stHeading"] h1 {

        display: inline-block;

        font-size: 32px !important;

        font-weight: 800 !important;

        color:
            #17336f !important;

        letter-spacing:
            -0.9px;

        margin-top:
            0 !important;

        margin-bottom:
            0 !important;

        padding-bottom:
            9px !important;

        text-shadow:

            0 3px 8px
            rgba(49,78,160,0.12),

            0 10px 25px
            rgba(49,78,160,0.07);

        transition:
            transform 0.35s ease;
    }


    [data-testid="stHeading"] h1:hover {

        transform:
            translateY(-3px)
            scale(1.01);
    }


    /* ========================================================
       TITLE UNDERLINE
       ======================================================== */

    [data-testid="stHeading"] h1::after {

        content: "";

        position: absolute;

        left: 50%;

        bottom: 0;

        transform:
            translateX(-50%);

        width: 85px;

        height: 5px;

        border-radius:
            10px;

        background:
            linear-gradient(
                90deg,
                #4f46e5,
                #06b6d4,
                #6366f1
            );

        box-shadow:
            0 3px 10px
            rgba(79,70,229,0.25);

        animation:
            titlePulse
            3s ease-in-out infinite;
    }


    @keyframes titlePulse {

        0% {

            width: 60px;
        }

        50% {

            width: 110px;
        }

        100% {

            width: 60px;
        }
    }


    /* ========================================================
       CAPTION
       ======================================================== */

    [data-testid="stCaptionContainer"] {

        text-align: center !important;

        color:
            #596b8c !important;

        font-size:
            12px !important;

        font-weight:
            600 !important;
    }


    /* ========================================================
       NAVIGATION BUTTONS
       ======================================================== */

    div[data-testid="stHorizontalBlock"] button {

        min-height:
            40px !important;

        border-radius:
            10px !important;

        font-weight:
            650 !important;

        border:
            1px solid
            rgba(185,196,220,0.9) !important;

        background:
            linear-gradient(
                145deg,
                #ffffff,
                #f5f7fc
            ) !important;

        color:
            #294277 !important;

        box-shadow:

            0 5px 0
            rgba(74,92,145,0.07),

            0 8px 18px
            rgba(55,78,140,0.08);

        transition:
            transform 0.2s ease,
            box-shadow 0.2s ease;
    }


    div[data-testid="stHorizontalBlock"] button:hover {

        transform:
            translateY(-3px);

        box-shadow:

            0 8px 0
            rgba(74,92,145,0.07),

            0 14px 25px
            rgba(55,78,140,0.14);
    }


    /* ========================================================
       SELECTED PAGE = LIGHT ORANGE
       ======================================================== */

    div[data-testid="stHorizontalBlock"]
    button[kind="primary"] {

        background:
            linear-gradient(
                145deg,
                #fff4df,
                #ffe8c2
            ) !important;

        color:
            #9a5b00 !important;

        border:
            1px solid
            #f3c77c !important;

        box-shadow:

            0 5px 0
            rgba(225,157,67,0.08),

            0 8px 18px
            rgba(225,157,67,0.12);
    }


    div[data-testid="stHorizontalBlock"]
    button[kind="primary"]:hover {

        background:
            linear-gradient(
                145deg,
                #ffefd4,
                #ffdfae
            ) !important;

        color:
            #8d5100 !important;

        border-color:
            #eeb866 !important;
    }


    /* ========================================================
       DIVIDER
       ======================================================== */

    hr {

        border: none !important;

        border-top:
            1px solid
            #e4e7ec !important;

        margin-top:
            9px !important;

        margin-bottom:
            13px !important;
    }


    /* ========================================================
       FILTER LABELS
       ======================================================== */

    .filter-title {

        font-size:
            12px;

        font-weight:
            700;

        color:
            #344054;

        margin-bottom:
            4px;
    }


    /* ========================================================
       FILTER BUTTONS
       ======================================================== */

    div[data-testid="stPopover"] > button {

        width:
            100% !important;

        min-height:
            40px !important;

        border-radius:
            9px !important;

        border:
            1px solid
            #d0d5dd !important;

        background:
            linear-gradient(
                145deg,
                #ffffff,
                #fafbff
            ) !important;

        color:
            #344054 !important;

        font-size:
            13px !important;

        font-weight:
            550 !important;

        box-shadow:
            0 3px 10px
            rgba(16,24,40,0.04);

        transition:
            transform 0.25s ease,
            border-color 0.25s ease,
            box-shadow 0.25s ease;
    }


    div[data-testid="stPopover"] > button:hover {

        transform:
            translateY(-2px);

        border-color:
            #7c83ff !important;

        box-shadow:
            0 8px 18px
            rgba(79,70,229,0.10);
    }


    /* ========================================================
       KPI CARDS
       ======================================================== */

    [data-testid="stMetric"] {

        background:
            linear-gradient(
                145deg,
                rgba(255,255,255,0.98),
                rgba(247,249,255,0.96)
            ) !important;

        border:
            1px solid
            rgba(211,220,240,0.95) !important;

        border-radius:
            13px !important;

        padding:
            13px 13px 17px 13px !important;

        min-height:
            75px !important;

        position:
            relative;

        overflow:
            hidden;

        box-shadow:

            0 4px 10px
            rgba(37,60,120,0.05),

            0 15px 30px
            rgba(37,60,120,0.04);

        transform-style:
            preserve-3d;

        transition:
            transform 0.35s ease,
            box-shadow 0.35s ease;
    }


    [data-testid="stMetric"]:hover {

        transform:
            translateY(-6px)
            rotateX(3deg)
            rotateY(-2deg);

        box-shadow:

            0 10px 20px
            rgba(37,60,120,0.09),

            0 22px 40px
            rgba(37,60,120,0.10);
    }


    [data-testid="stMetric"]::before {

        content: "";

        position:
            absolute;

        width:
            75px;

        height:
            75px;

        right:
            -30px;

        top:
            -30px;

        border-radius:
            50%;

        background:
            radial-gradient(
                circle,
                rgba(99,102,241,0.11),
                transparent 68%
            );

        animation:
            metricGlow
            4s ease-in-out infinite;
    }


    @keyframes metricGlow {

        0%,
        100% {

            transform:
                translate3d(0,0,0)
                scale(1);
        }

        50% {

            transform:
                translate3d(-12px,12px,0)
                scale(1.25);
        }
    }


    /* ========================================================
       KPI BOTTOM LINE
       ======================================================== */

    [data-testid="stMetric"]::after {

        content: "";

        position:
            absolute;

        bottom:
            0;

        left:
            0;

        right:
            0;

        height:
            4px;

        background:
            linear-gradient(
                90deg,
                #3b82f6,
                #06b6d4,
                #6366f1
            );

        background-size:
            200% 100%;

        animation:
            gradientMove
            4s linear infinite;
    }


    @keyframes gradientMove {

        0% {

            background-position:
                0% 50%;
        }

        100% {

            background-position:
                200% 50%;
        }
    }


    [data-testid="stMetricLabel"] {

        color:
            #65738b !important;

        font-size:
            12px !important;

        font-weight:
            600 !important;
    }


    [data-testid="stMetricValue"] {

        color:
            #17336f !important;

        font-size:
            27px !important;

        font-weight:
            800 !important;

        text-shadow:
            0 2px 7px
            rgba(36,62,125,0.10);
    }


    /* ========================================================
       SECTION HEADINGS
       ======================================================== */

    .section-heading {

        color:
            #17336f;

        font-size:
            17px;

        font-weight:
            800;

        border-left:
            4px solid
            #4f46e5;

        padding-left:
            11px;

        margin-top:
            12px;

        margin-bottom:
            8px;

        position:
            relative;
    }


    .section-heading::after {

        content: "";

        display:
            inline-block;

        width:
            25px;

        height:
            3px;

        margin-left:
            8px;

        vertical-align:
            middle;

        border-radius:
            4px;

        background:
            linear-gradient(
                90deg,
                #06b6d4,
                #6366f1
            );

        animation:
            sectionPulse
            2.5s ease-in-out infinite;
    }


    @keyframes sectionPulse {

        0% {

            width:
                18px;

            opacity:
                0.5;
        }

        50% {

            width:
                35px;

            opacity:
                1;
        }

        100% {

            width:
                18px;

            opacity:
                0.5;
        }
    }


    /* ========================================================
       REFRESH TEXT
       ======================================================== */

    .refresh-text {

        color:
            #8993a5;

        font-size:
            11px;

        text-align:
            center;

        margin-top:
            2px;

        margin-bottom:
            5px;
    }


    /* ========================================================
       CHART
       ======================================================== */

    [data-testid="stPlotlyChart"] {

        border-radius:
            12px;

        transition:
            transform 0.35s ease,
            filter 0.35s ease;
    }


    [data-testid="stPlotlyChart"]:hover {

        transform:
            translateY(-3px);

        filter:
            drop-shadow(
                0 10px 16px
                rgba(52,73,140,0.08)
            );
    }


    /* ========================================================
       TABLE
       ======================================================== */

    [data-testid="stDataFrame"] {

        border:
            1px solid
            #e2e6ef !important;

        border-radius:
            9px !important;

        overflow:
            hidden !important;

        box-shadow:
            0 4px 15px
            rgba(16,24,40,0.05);

        transition:
            box-shadow 0.25s ease,
            transform 0.25s ease;
    }


    [data-testid="stDataFrame"]:hover {

        transform:
            translateY(-2px);

        box-shadow:
            0 8px 22px
            rgba(16,24,40,0.08);
    }


    /* ========================================================
       DOWNLOAD BUTTON
       ======================================================== */

    [data-testid="stDownloadButton"] button {

        border-radius:
            8px !important;

        font-weight:
            650 !important;

        transition:
            transform 0.2s ease,
            box-shadow 0.2s ease;
    }


    [data-testid="stDownloadButton"] button:hover {

        transform:
            translateY(-2px);

        box-shadow:
            0 8px 16px
            rgba(52,73,140,0.10);
    }


    /* ========================================================
       MOBILE
       ======================================================== */

    @media (max-width: 800px) {

        .block-container {

            max-width:
                100% !important;

            padding-left:
                0.7rem !important;

            padding-right:
                0.7rem !important;
        }


        [data-testid="stHeading"] h1 {

            font-size:
                24px !important;
        }


        [data-testid="stMetricValue"] {

            font-size:
                22px !important;
        }


        [data-testid="stApp"] .main::before {

            font-size:
                75px !important;

            letter-spacing:
                10px !important;

            width:
                180vw !important;
        }
    }


    /* ========================================================
       REDUCED MOTION
       ======================================================== */

    @media (prefers-reduced-motion: reduce) {

        *,
        *::before,
        *::after {

            animation-duration:
                0.01ms !important;

            animation-iteration-count:
                1 !important;

            transition-duration:
                0.01ms !important;
        }
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel_data():

    if not os.path.exists(EXCEL_FILE):

        raise FileNotFoundError(
            f"{EXCEL_FILE} was not found "
            "in the same folder as app.py."
        )


    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name="Sheet1",
        engine="openpyxl"
    )


    # --------------------------------------------------------
    # CLEAN COLUMN NAMES
    # --------------------------------------------------------

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )


    # --------------------------------------------------------
    # REQUIRED COLUMNS
    # --------------------------------------------------------

    required_columns = [

        "Competitor",
        "Evidence",
        "MarketTopic",
        "PublicationDate",
        "SourceURL",
        "Status",
        "Summary",
        "Title"
    ]


    missing_columns = [

        column

        for column in required_columns

        if column not in df.columns
    ]


    if missing_columns:

        raise ValueError(
            "Missing required columns: "
            +
            ", ".join(
                missing_columns
            )
        )


    # --------------------------------------------------------
    # CLEAN TEXT
    # --------------------------------------------------------

    text_columns = [

        "Competitor",
        "Evidence",
        "MarketTopic",
        "SourceURL",
        "Status",
        "Summary",
        "Title"
    ]


    for column in text_columns:

        df[column] = (

            df[column]

            .fillna("")

            .astype(str)

            .str.strip()
        )


    # --------------------------------------------------------
    # BLANK STATUS
    # --------------------------------------------------------

    df["Status"] = (

        df["Status"]

        .replace(
            "",
            STATUS_NOT_PROVIDED
        )
    )


    # --------------------------------------------------------
    # CLEAN URL
    # --------------------------------------------------------

    def clean_url(value):

        value = str(
            value
        ).strip()


        if not value:

            return ""


        value = value.replace(
            "\\",
            ""
        )


        match = re.search(
            r"https?://[^\s\]\)]+",
            value
        )


        if match:

            return match.group(0)


        return value


    df["SourceURL"] = (

        df["SourceURL"]

        .apply(
            clean_url
        )
    )


    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    df["PublicationDate"] = (

        pd.to_datetime(
            df["PublicationDate"],
            errors="coerce"
        )
    )


    return df


# ============================================================
# DATA VERSION
# ============================================================

def get_data_version(df):

    try:

        data_hash = (

            pd.util.hash_pandas_object(
                df.fillna("").astype(str),
                index=True
            )
            .values
            .tobytes()
        )


        file_time = os.path.getmtime(
            EXCEL_FILE
        )


        raw = (

            str(file_time)

            +

            hashlib.md5(
                data_hash
            ).hexdigest()
        )


        return hashlib.md5(
            raw.encode()
        ).hexdigest()[:10]


    except Exception:

        return str(
            datetime.now().timestamp()
        )


# ============================================================
# FILTER CALLBACKS
# ============================================================

def select_all_changed(
    all_key,
    item_keys
):

    select_all = st.session_state.get(
        all_key,
        True
    )


    for key in item_keys:

        st.session_state[key] = (
            select_all
        )


def individual_changed(
    all_key,
    item_keys
):

    if not item_keys:

        st.session_state[
            all_key
        ] = True

        return


    all_selected = all(

        st.session_state.get(
            key,
            False
        )

        for key in item_keys
    )


    st.session_state[
        all_key
    ] = all_selected


# ============================================================
# CHECKBOX DROPDOWN
# ============================================================

def checkbox_dropdown(
    label,
    options,
    prefix,
    version
):

    options = list(
        options
    )


    all_key = (
        f"{prefix}_all_{version}"
    )


    item_keys = [

        f"{prefix}_{version}_{i}"

        for i in range(
            len(options)
        )
    ]


    # --------------------------------------------------------
    # INITIALIZE
    # --------------------------------------------------------

    if all_key not in st.session_state:

        st.session_state[
            all_key
        ] = True


    for key in item_keys:

        if key not in st.session_state:

            st.session_state[
                key
            ] = True


    # --------------------------------------------------------
    # LABEL
    # --------------------------------------------------------

    st.markdown(
        f"**{label}**"
    )


    # --------------------------------------------------------
    # SELECTED OPTIONS
    # --------------------------------------------------------

    selected = [

        option

        for option, key

        in zip(
            options,
            item_keys
        )

        if st.session_state.get(
            key,
            False
        )
    ]


    # --------------------------------------------------------
    # BUTTON TEXT
    # --------------------------------------------------------

    if len(selected) == len(options):

        if label == "Competitor":

            button_text = (
                "All Competitors ▾"
            )

        elif label == "Market Topic":

            button_text = (
                "All Market Topics ▾"
            )

        else:

            button_text = (
                "All Status ▾"
            )


    elif len(selected) == 0:

        button_text = (
            f"Select {label} ▾"
        )


    else:

        button_text = (
            f"{len(selected)} selected ▾"
        )


    # --------------------------------------------------------
    # POPOVER
    # --------------------------------------------------------

    with st.popover(
        button_text,
        use_container_width=True
    ):

        st.checkbox(
            "Select All",

            key=all_key,

            on_change=select_all_changed,

            args=(
                all_key,
                item_keys
            )
        )


        st.divider()


        for option, key in zip(
            options,
            item_keys
        ):

            st.checkbox(
                str(option),

                key=key,

                on_change=individual_changed,

                args=(
                    all_key,
                    item_keys
                )
            )


    # --------------------------------------------------------
    # FINAL SELECTION
    # --------------------------------------------------------

    return [

        option

        for option, key

        in zip(
            options,
            item_keys
        )

        if st.session_state.get(
            key,
            False
        )
    ]


# ============================================================
# HEADER
# ============================================================

st.title(
    "📊 Competitive Intelligence Dashboard"
)


st.caption(
    "Monitor competitor activities, market topics and strategic intelligence"
)


# ============================================================
# NAVIGATION
# ============================================================

nav1, nav2 = st.columns(
    [1, 1.15]
)


with nav1:

    if (
        st.session_state.page
        == "Executive Overview"
    ):

        if st.button(
            "📊 Overview",
            type="primary",
            use_container_width=True,
            key="overview_button"
        ):

            st.session_state.page = (
                "Executive Overview"
            )

            st.rerun()

    else:

        if st.button(
            "📊 Overview",
            use_container_width=True,
            key="overview_button"
        ):

            st.session_state.page = (
                "Executive Overview"
            )

            st.rerun()


with nav2:

    if (
        st.session_state.page
        == "Intelligence Details"
    ):

        if st.button(
            "📋 Intelligence Details",
            type="primary",
            use_container_width=True,
            key="details_button"
        ):

            st.session_state.page = (
                "Intelligence Details"
            )

            st.rerun()

    else:

        if st.button(
            "📋 Intelligence Details",
            use_container_width=True,
            key="details_button"
        ):

            st.session_state.page = (
                "Intelligence Details"
            )

            st.rerun()


# ============================================================
# DIVIDER
# ============================================================

st.divider()


# ============================================================
# MAIN DASHBOARD
# ============================================================

@st.fragment(run_every="5s")
def dashboard():

    # --------------------------------------------------------
    # LOAD DATA
    # --------------------------------------------------------

    try:

        df = load_excel_data()

    except Exception as error:

        st.error(
            "Unable to load output.xlsx."
        )

        st.error(
            str(error)
        )

        return


    # --------------------------------------------------------
    # DATA VERSION
    # --------------------------------------------------------

    version = get_data_version(
        df
    )


    # ========================================================
    # FILTER OPTIONS
    # ========================================================

    competitors = sorted(

        [

            value

            for value in df[
                "Competitor"
            ].unique()

            if str(value).strip()
        ],

        key=str.lower
    )


    market_topics = sorted(

        [

            value

            for value in df[
                "MarketTopic"
            ].unique()

            if str(value).strip()
        ],

        key=str.lower
    )


    statuses = sorted(

        [

            value

            for value in df[
                "Status"
            ].unique()

            if str(value).strip()
        ],

        key=str.lower
    )


    # ========================================================
    # FILTER ROW
    # ========================================================

    filter1, filter2, filter3 = st.columns(
        3
    )


    with filter1:

        selected_competitors = (
            checkbox_dropdown(
                "Competitor",
                competitors,
                "competitor",
                version
            )
        )


    with filter2:

        selected_topics = (
            checkbox_dropdown(
                "Market Topic",
                market_topics,
                "market_topic",
                version
            )
        )


    with filter3:

        selected_statuses = (
            checkbox_dropdown(
                "Status",
                statuses,
                "status",
                version
            )
        )


    # ========================================================
    # APPLY FILTERS
    # ========================================================

    filtered_df = df.copy()


    if selected_competitors:

        filtered_df = filtered_df[
            filtered_df[
                "Competitor"
            ].isin(
                selected_competitors
            )
        ]

    else:

        filtered_df = (
            filtered_df.iloc[0:0]
        )


    if selected_topics:

        filtered_df = filtered_df[
            filtered_df[
                "MarketTopic"
            ].isin(
                selected_topics
            )
        ]

    else:

        filtered_df = (
            filtered_df.iloc[0:0]
        )


    if selected_statuses:

        filtered_df = filtered_df[
            filtered_df[
                "Status"
            ].isin(
                selected_statuses
            )
        ]

    else:

        filtered_df = (
            filtered_df.iloc[0:0]
        )


    # ========================================================
    # REFRESH MESSAGE
    # ========================================================

    st.caption(
        (
            f"Showing {len(filtered_df)} "
            f"of {len(df)} Excel records "
            f"• Last refreshed: "
            f"{datetime.now().strftime('%d-%b-%Y %H:%M:%S')}"
        )
    )


    # ========================================================
    # EXECUTIVE OVERVIEW
    # ========================================================

    if (
        st.session_state.page
        == "Executive Overview"
    ):

        # ----------------------------------------------------
        # KPI CALCULATIONS
        # ----------------------------------------------------

        total_insights = len(
            filtered_df
        )


        companies_tracked = (

            filtered_df[
                "Competitor"
            ]

            .replace(
                "",
                pd.NA
            )

            .dropna()

            .nunique()
        )


        relevant_information = (

            filtered_df[
                "Status"
            ]

            .astype(str)

            .str.casefold()

            .eq(
                STATUS_RELEVANT.casefold()
            )

            .sum()
        )


        topics_covered = (

            filtered_df[
                "MarketTopic"
            ]

            .replace(
                "",
                pd.NA
            )

            .dropna()

            .nunique()
        )


        # ----------------------------------------------------
        # KPI CARDS
        # ----------------------------------------------------

        k1, k2, k3, k4 = st.columns(
            4
        )


        with k1:

            st.metric(
                "Total Insights",
                total_insights
            )


        with k2:

            st.metric(
                "Companies Tracked",
                companies_tracked
            )


        with k3:

            st.metric(
                "Relevant Information",
                relevant_information
            )


        with k4:

            st.metric(
                "Topics Covered",
                topics_covered
            )


        # ====================================================
        # SECTION
        # ====================================================

        st.markdown(
            "### Competitive Intelligence Activity"
        )


        # ====================================================
        # CHART COLUMNS
        # ====================================================

        chart1, chart2 = st.columns(
            2
        )


        # ====================================================
        # COMPETITOR CHART
        # ====================================================

        with chart1:

            company_counts = (

                filtered_df

                .groupby(
                    "Competitor"
                )

                .size()

                .reset_index(
                    name="Insights"
                )

                .sort_values(
                    "Insights",
                    ascending=True
                )
            )


            if not company_counts.empty:

                fig_company = px.bar(

                    company_counts,

                    x="Insights",

                    y="Competitor",

                    orientation="h",

                    text="Insights",

                    template="plotly_white"
                )


                fig_company.update_traces(

                    textposition="outside",

                    cliponaxis=False,

                    marker_color="#8DB0E8",

                    marker_line_width=0
                )


                fig_company.update_layout(

                    height=360,

                    margin=dict(
                        l=5,
                        r=35,
                        t=5,
                        b=35
                    ),

                    xaxis_title=
                        "Number of Insights",

                    yaxis_title="",

                    showlegend=False,

                    plot_bgcolor=
                        "rgba(255,255,255,0)",

                    paper_bgcolor=
                        "rgba(255,255,255,0)",

                    font=dict(
                        color="#7A8498"
                    ),

                    hoverlabel=dict(
                        bgcolor="#17336f",
                        font_color="white"
                    )
                )


                st.plotly_chart(

                    fig_company,

                    use_container_width=True,

                    key=(
                        "company_chart_"
                        + version
                    )
                )


            else:

                st.info(
                    "No competitor data available."
                )


        # ====================================================
        # MARKET TOPIC CHART
        # ====================================================

        with chart2:

            topic_counts = (

                filtered_df

                .groupby(
                    "MarketTopic"
                )

                .size()

                .reset_index(
                    name="Insights"
                )

                .sort_values(
                    "Insights",
                    ascending=True
                )
            )


            if not topic_counts.empty:

                fig_topic = px.bar(

                    topic_counts,

                    x="Insights",

                    y="MarketTopic",

                    orientation="h",

                    text="Insights",

                    template="plotly_white"
                )


                fig_topic.update_traces(

                    textposition="outside",

                    cliponaxis=False,

                    marker_color="#B3A0E5",

                    marker_line_width=0
                )


                fig_topic.update_layout(

                    height=360,

                    margin=dict(
                        l=5,
                        r=35,
                        t=5,
                        b=35
                    ),

                    xaxis_title=
                        "Number of Insights",

                    yaxis_title="",

                    showlegend=False,

                    plot_bgcolor=
                        "rgba(255,255,255,0)",

                    paper_bgcolor=
                        "rgba(255,255,255,0)",

                    font=dict(
                        color="#7A8498"
                    ),

                    hoverlabel=dict(
                        bgcolor="#17336f",
                        font_color="white"
                    )
                )


                st.plotly_chart(

                    fig_topic,

                    use_container_width=True,

                    key=(
                        "topic_chart_"
                        + version
                    )
                )


            else:

                st.info(
                    "No market-topic data available."
                )


        # ====================================================
        # COMPANY VS MARKET TOPIC
        # ====================================================

        st.markdown(
            "### Company vs Market Topic"
        )


        if not filtered_df.empty:

            matrix = pd.crosstab(

                filtered_df[
                    "MarketTopic"
                ],

                filtered_df[
                    "Competitor"
                ]
            )


            if selected_competitors:

                matrix = matrix.reindex(

                    columns=
                        selected_competitors,

                    fill_value=0
                )


            matrix = matrix.sort_index()

            matrix = matrix.reset_index()


            if not matrix.empty:

                matrix_height = min(

                    max(
                        180,
                        37 * (
                            len(matrix) + 1
                        )
                    ),

                    520
                )


                st.dataframe(

                    matrix,

                    use_container_width=True,

                    hide_index=True,

                    height=matrix_height,

                    key=(
                        "matrix_"
                        + version
                    )
                )


            else:

                st.info(
                    "No matrix data available."
                )


        else:

            st.info(
                "No data available for "
                "the selected filters."
            )


    # ========================================================
    # INTELLIGENCE DETAILS
    # ========================================================

    else:

        st.markdown(
            "### Intelligence Details"
        )


        display_df = (
            filtered_df.copy()
        )


        display_df[
            "PublicationDate"
        ] = (

            display_df[
                "PublicationDate"
            ]

            .dt.strftime(
                "%d-%b-%Y"
            )

            .fillna(
                "Not available"
            )
        )


        display_columns = [

            "Competitor",
            "MarketTopic",
            "Title",
            "PublicationDate",
            "Summary",
            "Status",
            "Evidence",
            "SourceURL"
        ]


        display_columns = [

            column

            for column in
            display_columns

            if column in
            display_df.columns
        ]


        display_df = display_df[
            display_columns
        ]


        if not display_df.empty:

            table_height = min(

                max(
                    220,
                    42 * len(
                        display_df
                    ) + 55
                ),

                650
            )


            st.dataframe(

                display_df,

                use_container_width=True,

                hide_index=True,

                height=table_height,

                column_config={

                    "SourceURL":

                        st.column_config.LinkColumn(

                            "Source URL",

                            display_text=
                                "Open Source"
                        )
                },

                key=(
                    "details_"
                    + version
                )
            )


        else:

            st.info(
                "No intelligence records "
                "match the selected filters."
            )


        # ----------------------------------------------------
        # DOWNLOAD
        # ----------------------------------------------------

        csv_data = (

            display_df

            .to_csv(
                index=False
            )

            .encode(
                "utf-8"
            )
        )


        st.download_button(

            "⬇ Download Filtered Intelligence",

            data=csv_data,

            file_name=
                "filtered_intelligence.csv",

            mime="text/csv"
        )


# ============================================================
# RUN
# ============================================================

dashboard()