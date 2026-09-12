from __future__ import annotations

from pathlib import Path


APP_VERSION = "0.2"
COMPANY_NAME = "JAY TEA"
PRODUCT_NAME = "Order Planning Studio"
PREMIUM_BADGE = "Precision planning • Local & private"
VALUE_PROPOSITION = (
    "Turn the daily ERP Order Status workbook into a clear, validated planning view "
    "for the first 36 eligible orders—in one polished Excel report."
)

APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "assets" / "logo.png"

# Palette extracted from the supplied JAY logo.
GOLD = "#E0A010"
GOLD_BRIGHT = "#F0D15B"
GOLD_SOFT = "#F0E080"
OBSIDIAN = "#08090B"
CHARCOAL = "#111317"
WARM_WHITE = "#F7F3E8"
MUTED_SILVER = "#B8B5AD"


def liquid_glass_css() -> str:
    """Return the centralized visual system for the Streamlit interface."""
    return f"""
<style>
:root {{
  --brand-gold: {GOLD};
  --brand-gold-bright: {GOLD_BRIGHT};
  --brand-gold-soft: {GOLD_SOFT};
  --brand-obsidian: {OBSIDIAN};
  --brand-charcoal: {CHARCOAL};
  --brand-white: {WARM_WHITE};
  --brand-muted: {MUTED_SILVER};
  --glass-border: rgba(255, 255, 255, 0.13);
  --glass-highlight: rgba(255, 255, 255, 0.08);
  --glass-bg: rgba(20, 22, 26, 0.66);
  --glass-bg-strong: rgba(18, 20, 24, 0.82);
  --glass-shadow: 0 24px 70px rgba(0, 0, 0, 0.38), 0 1px 0 rgba(255,255,255,.06) inset;
}}

html {{ color-scheme: dark; }}

.stApp {{
  background:
    radial-gradient(circle at 12% 8%, rgba(224, 160, 16, 0.18), transparent 26rem),
    radial-gradient(circle at 88% 22%, rgba(240, 209, 91, 0.09), transparent 30rem),
    linear-gradient(145deg, #07080a 0%, #0d0f12 48%, #090a0c 100%);
  color: var(--brand-white);
}}

.stApp::before {{
  content: "";
  position: fixed;
  inset: 0;
  pointer-events: none;
  background-image: linear-gradient(rgba(255,255,255,.012) 1px, transparent 1px),
                    linear-gradient(90deg, rgba(255,255,255,.012) 1px, transparent 1px);
  background-size: 42px 42px;
  mask-image: linear-gradient(to bottom, black, transparent 78%);
}}

header[data-testid="stHeader"] {{
  background: rgba(8, 9, 11, 0.56);
  backdrop-filter: blur(24px) saturate(155%);
  border-bottom: 1px solid rgba(255,255,255,.06);
}}

.stMainBlockContainer,
[data-testid="stMainBlockContainer"] {{
  max-width: 1240px;
  padding-top: 4.8rem !important;
  padding-bottom: 5rem;
}}

.st-key-brand_nav {{
  position: relative;
  z-index: 99;
  padding: .65rem 1rem;
  margin-bottom: 1.25rem;
  border: 1px solid var(--glass-border);
  border-radius: 999px;
  background: rgba(18, 20, 24, .72);
  backdrop-filter: blur(28px) saturate(165%);
  box-shadow: 0 16px 42px rgba(0,0,0,.26), 0 1px 0 rgba(255,255,255,.08) inset;
}}

.st-key-brand_nav img {{
  filter: drop-shadow(0 5px 14px rgba(224,160,16,.22));
}}

.st-key-brand_nav p {{
  margin: 0;
  line-height: 1;
  white-space: nowrap;
  display: flex;
  align-items: center;
}}

.st-key-brand_nav_inner {{
  width: 100%;
}}

.st-key-brand_nav_inner [data-testid="stImage"],
.st-key-brand_nav_inner [data-testid="stMarkdownContainer"] {{
  display: flex;
  align-items: center;
  margin: 0;
}}

.st-key-brand_nav_inner [data-testid="stImage"] figure {{
  margin: 0;
}}

.st-key-hero {{
  position: relative;
  overflow: hidden;
  padding: clamp(1.5rem, 4vw, 3.2rem);
  margin-bottom: 1.5rem;
  border: 1px solid rgba(240, 209, 91, .25);
  border-radius: 32px;
  background:
    linear-gradient(135deg, rgba(255,255,255,.10), rgba(255,255,255,.025) 44%, rgba(224,160,16,.08)),
    rgba(17, 19, 23, .70);
  backdrop-filter: blur(30px) saturate(160%);
  box-shadow: var(--glass-shadow), 0 0 90px rgba(224,160,16,.06);
}}

.st-key-hero::before {{
  content: "";
  position: absolute;
  width: 24rem;
  height: 24rem;
  right: -8rem;
  top: -13rem;
  border-radius: 50%;
  background: radial-gradient(circle, rgba(240,209,91,.21), transparent 68%);
  pointer-events: none;
}}

.st-key-hero img {{
  filter: drop-shadow(0 22px 40px rgba(0,0,0,.52)) drop-shadow(0 0 24px rgba(224,160,16,.12));
  transition: transform .28s ease, filter .28s ease;
}}

.st-key-hero img:hover {{
  transform: translateY(-3px) scale(1.015);
  filter: drop-shadow(0 26px 46px rgba(0,0,0,.58)) drop-shadow(0 0 30px rgba(224,160,16,.18));
}}

.st-key-hero h1 {{
  margin: .35rem 0 .55rem;
  letter-spacing: -.045em;
  line-height: 1.02;
  background: linear-gradient(100deg, var(--brand-white) 12%, #fff 50%, var(--brand-gold-soft));
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}}

.st-key-hero p {{
  max-width: 48rem;
  color: #D2CEC4;
  font-size: 1.05rem;
  line-height: 1.65;
}}

.st-key-workflow_intro {{ margin: 2.1rem 0 .8rem; }}
.st-key-workflow_intro h2 {{ letter-spacing: -.025em; }}

[data-testid="stVerticalBlockBorderWrapper"] {{
  background: linear-gradient(145deg, rgba(255,255,255,.075), rgba(255,255,255,.025)), var(--glass-bg);
  backdrop-filter: blur(24px) saturate(150%);
  border: 1px solid var(--glass-border);
  border-radius: 24px;
  box-shadow: 0 18px 50px rgba(0,0,0,.24), 0 1px 0 rgba(255,255,255,.06) inset;
  transition: transform .22s ease, border-color .22s ease, box-shadow .22s ease;
}}

[data-testid="stVerticalBlockBorderWrapper"]:hover {{
  transform: translateY(-2px);
  border-color: rgba(240,209,91,.28);
  box-shadow: 0 22px 58px rgba(0,0,0,.3), 0 0 0 1px rgba(224,160,16,.04) inset;
}}

.st-key-upload_panel, .st-key-output_panel {{ min-height: 22.5rem; }}

section[data-testid="stFileUploaderDropzone"] {{
  min-height: 10rem;
  border: 1px dashed rgba(240,209,91,.42);
  border-radius: 20px;
  background: linear-gradient(145deg, rgba(224,160,16,.09), rgba(255,255,255,.025));
  transition: border-color .2s ease, background .2s ease, transform .2s ease;
}}

section[data-testid="stFileUploaderDropzone"]:hover {{
  border-color: var(--brand-gold-bright);
  background: linear-gradient(145deg, rgba(224,160,16,.14), rgba(255,255,255,.035));
  transform: translateY(-1px);
}}

button, [role="button"], input, [role="radio"] {{
  transition: border-color .18s ease, box-shadow .18s ease, transform .18s ease, background .18s ease;
}}

.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {{
  border: 1px solid rgba(255,231,137,.72);
  border-radius: 999px;
  background: linear-gradient(135deg, #F5D45D 0%, #E0A010 54%, #B87308 100%);
  color: #15110A;
  font-weight: 750;
  box-shadow: 0 12px 30px rgba(224,160,16,.24), 0 1px 0 rgba(255,255,255,.5) inset;
}}

.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {{
  transform: translateY(-2px);
  border-color: #FFF0A5;
  box-shadow: 0 16px 38px rgba(224,160,16,.34), 0 1px 0 rgba(255,255,255,.58) inset;
}}

.stButton > button:focus-visible,
.stDownloadButton > button:focus-visible,
input:focus-visible,
[role="radio"]:focus-visible {{
  outline: 2px solid var(--brand-gold-bright);
  outline-offset: 3px;
}}

[data-baseweb="button-group"] {{
  padding: .3rem;
  border: 1px solid rgba(255,255,255,.09);
  border-radius: 18px;
  background: rgba(6,7,9,.42);
}}

[data-baseweb="button-group"] button {{
  border-radius: 13px !important;
  color: #D8D4CA;
}}

[data-baseweb="button-group"] button[aria-checked="true"] {{
  color: #171109;
  background: linear-gradient(135deg, var(--brand-gold-soft), var(--brand-gold));
  box-shadow: 0 7px 18px rgba(224,160,16,.18);
}}

[data-testid="stMetric"] {{
  min-width: 10rem;
  padding: 1.05rem 1.15rem;
  border-radius: 18px;
  background: linear-gradient(145deg, rgba(255,255,255,.07), rgba(255,255,255,.025));
  border: 1px solid rgba(255,255,255,.10);
}}

[data-testid="stMetricValue"] {{ color: var(--brand-gold-soft); }}

[data-testid="stTabs"] [role="tablist"] {{
  gap: .4rem;
  padding: .35rem;
  border-radius: 17px;
  background: rgba(7,8,10,.46);
  border: 1px solid rgba(255,255,255,.08);
}}

[data-testid="stTabs"] button[role="tab"] {{ border-radius: 12px; }}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
  background: rgba(224,160,16,.14);
  color: var(--brand-gold-soft);
}}

[data-testid="stDataFrame"] {{
  overflow: hidden;
  border: 1px solid rgba(255,255,255,.10);
  border-radius: 18px;
  background: rgba(8,9,11,.38);
}}

[data-testid="stAlert"] {{
  border-radius: 18px;
  backdrop-filter: blur(18px);
  border: 1px solid rgba(255,255,255,.13);
}}

[data-testid="stExpander"] {{
  margin-top: 1.2rem;
  border: 1px solid rgba(255,255,255,.10);
  border-radius: 20px;
  overflow: hidden;
  background: rgba(17,19,23,.58);
  backdrop-filter: blur(20px);
}}

.st-key-action_bar {{
  padding: 1rem 1.1rem;
  margin: .35rem 0 1.8rem;
  border-radius: 22px;
  background: rgba(17,19,23,.66);
  border: 1px solid rgba(255,255,255,.10);
  backdrop-filter: blur(22px) saturate(150%);
  box-shadow: 0 15px 40px rgba(0,0,0,.22);
}}

.st-key-result_shell {{
  padding: clamp(1rem, 3vw, 1.8rem);
  border-radius: 28px;
  border: 1px solid rgba(240,209,91,.20);
  background: rgba(15,17,20,.7);
  backdrop-filter: blur(26px) saturate(155%);
  box-shadow: var(--glass-shadow);
}}

small, [data-testid="stCaptionContainer"], .stCaption {{ color: var(--brand-muted) !important; }}
a {{ color: var(--brand-gold-soft); }}

@media (max-width: 768px) {{
  .stMainBlockContainer,
  [data-testid="stMainBlockContainer"] {{ padding: 4.4rem .85rem 3rem !important; }}
  .st-key-brand_nav {{ border-radius: 22px; padding: .55rem .7rem; }}
  .st-key-brand_nav [data-testid="stHorizontalBlock"] {{ gap: .5rem; }}
  .st-key-hero {{ border-radius: 24px; text-align: center; }}
  .st-key-hero img {{ max-width: 160px; margin-inline: auto; }}
  .st-key-hero h1 {{ font-size: 2.35rem; }}
  .st-key-upload_panel, .st-key-output_panel {{ min-height: auto; }}
  [data-baseweb="button-group"] {{ overflow-x: auto; justify-content: flex-start; }}
  [data-baseweb="button-group"] button {{ flex: 0 0 auto; }}
  [data-testid="stMetric"] {{ min-width: 100%; }}
  .stButton > button, .stDownloadButton > button {{ width: 100%; }}
}}

@media (prefers-reduced-motion: reduce) {{
  *, *::before, *::after {{ scroll-behavior: auto !important; transition-duration: .01ms !important; }}
}}
</style>
"""
