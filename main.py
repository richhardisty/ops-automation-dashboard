# ──────────────────────────────────────────────
# DEMO VERSION
# This is a sanitized portfolio copy of the app entry point.
# Credentials, internal file paths, and proprietary data have
# been removed. The individual page files referenced in the
# navigation are not included in this repo.
# See README.md for full context and architecture overview.
# ──────────────────────────────────────────────

import streamlit as st

# ──────────────────────────────────────────────
# App-wide settings
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Ops Automation Dashboard",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ──────────────────────────────────────────────
# Custom CSS — dark industrial theme
# ──────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Base ── */
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans', sans-serif;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: #0d0f14;
        border-right: 1px solid #1e2235;
    }
    [data-testid="stSidebar"] .stMarkdown p {
        color: #505568;
        font-size: 0.72rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    /* ── Main area ── */
    .main .block-container {
        padding-top: 2rem;
        max-width: 960px;
    }

    /* ── Hero header ── */
    .hero-header {
        border-left: 3px solid #40aaff;
        padding: 0.6rem 0 0.6rem 1.2rem;
        margin-bottom: 2rem;
    }
    .hero-header h1 {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.6rem;
        font-weight: 600;
        color: #e8eaf0;
        margin: 0 0 0.2rem 0;
        letter-spacing: -0.02em;
    }
    .hero-header p {
        color: #505568;
        font-size: 0.85rem;
        margin: 0;
    }

    /* ── Section label ── */
    .section-label {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.65rem;
        font-weight: 600;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: #40aaff;
        margin-bottom: 0.6rem;
        margin-top: 1.8rem;
    }

    /* ── Tool cards ── */
    .tool-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
        gap: 12px;
        margin-bottom: 1rem;
    }
    .tool-card {
        background: #0d0f14;
        border: 1px solid #1e2235;
        border-radius: 8px;
        padding: 16px 18px;
        transition: border-color 0.15s ease, background 0.15s ease;
        cursor: default;
    }
    .tool-card:hover {
        border-color: #40aaff55;
        background: #111520;
    }
    .tool-card .card-icon {
        font-size: 1.3rem;
        margin-bottom: 8px;
    }
    .tool-card .card-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
        font-weight: 600;
        color: #c8cce0;
        margin-bottom: 4px;
    }
    .tool-card .card-desc {
        font-size: 0.75rem;
        color: #505568;
        line-height: 1.5;
    }
    .tool-card .card-badge {
        display: inline-block;
        font-size: 0.62rem;
        font-family: 'IBM Plex Mono', monospace;
        padding: 2px 7px;
        border-radius: 4px;
        margin-top: 8px;
        letter-spacing: 0.06em;
        text-transform: uppercase;
    }
    .badge-live   { background: #0f3d26; color: #3ee87a; border: 1px solid #3ee87a44; }
    .badge-beta   { background: #2a1f0a; color: #f5c542; border: 1px solid #f5c54244; }
    .badge-stable { background: #0a2040; color: #40aaff; border: 1px solid #40aaff44; }

    /* ── Stats row ── */
    .stats-row {
        display: flex;
        gap: 12px;
        margin-bottom: 1.5rem;
        flex-wrap: wrap;
    }
    .stat-pill {
        background: #0d0f14;
        border: 1px solid #1e2235;
        border-radius: 6px;
        padding: 10px 16px;
        min-width: 120px;
    }
    .stat-pill .stat-val {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.3rem;
        font-weight: 600;
        color: #40aaff;
        line-height: 1.1;
    }
    .stat-pill .stat-lbl {
        font-size: 0.68rem;
        color: #505568;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-top: 2px;
    }

    /* ── Footer ── */
    .dash-footer {
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #1e2235;
        font-size: 0.72rem;
        color: #303448;
        font-family: 'IBM Plex Mono', monospace;
    }

    /* ── Streamlit overrides ── */
    .stInfo { background: #0a1a2e; border-color: #40aaff33; }
    div[data-testid="stDecoration"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding: 1rem 0 0.5rem 0;">
        <div style="font-family: 'IBM Plex Mono', monospace; font-size: 1rem;
                    font-weight: 600; color: #e8eaf0; letter-spacing: -0.01em;">
            ⚙️ Ops Dashboard
        </div>
        <div style="font-size: 0.7rem; color: #303448; margin-top: 4px;
                    font-family: 'IBM Plex Mono', monospace;">
            v1.0 · internal tooling
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

# ──────────────────────────────────────────────
# Tool definitions — sanitized for demo
# ──────────────────────────────────────────────
tools = {
    "Home": [
        st.Page(
            page="pages/home.py",
            title="Home",
            icon="🏠",
            default=True
        )
    ],
    "Analytics": [
        st.Page(
            "pages/Analytics/inventory_health_report.py",
            title="Inventory Health Report",
            icon="📊"
        ),
    ],
    "Amazon": [
        st.Page(
            "pages/Amazon/backorder_pipeline.py",
            title="Backorder Pipeline",
            icon="📦"
        ),
        st.Page(
            "pages/Amazon/generate_paperwork.py",
            title="Generate Order Paperwork",
            icon="📄"
        ),
        st.Page(
            "pages/Amazon/backorder_upload.py",
            title="Generate Backorder Upload",
            icon="⬆️"
        ),
        st.Page(
            "pages/Amazon/sales_report.py",
            title="Weekly Sales Report",
            icon="📈"
        ),
    ],
    "Snap-on": [
        st.Page(
            "pages/Snapon/pack_slip_generator.py",
            title="Pack Slip Generator",
            icon="🔧"
        ),
    ],
    "Walmart US": [
        st.Page(
            "pages/Walmart/pre_confirm_transfer.py",
            title="Pre-Confirm Transfer",
            icon="🛒"
        ),
        st.Page(
            "pages/Walmart/confirm_orders.py",
            title="Confirm Orders",
            icon="🛒"
        ),
        st.Page(
            "pages/Walmart/routed_orders_upload.py",
            title="Routed Orders Upload",
            icon="🛒"
        ),
        st.Page(
            "pages/Walmart/pack_slip.py",
            title="Generate Pack Slip",
            icon="🛒"
        ),
        st.Page(
            "pages/Walmart/carton_labels.py",
            title="Generate Carton Labels",
            icon="🛒"
        ),
        st.Page(
            "pages/Walmart/ltl_tl_documents.py",
            title="LTL & TL Labels + Documents",
            icon="🚚"
        ),
    ],
    "Walmart CA": [
        st.Page(
            "pages/WalmartCA/process_orders.py",
            title="Process Orders",
            icon="🛒"
        ),
        st.Page(
            "pages/WalmartCA/add_ship_data.py",
            title="Add Ship Data",
            icon="🛒"
        ),
    ],
}

# ──────────────────────────────────────────────
# Navigation
# ──────────────────────────────────────────────
pg = st.navigation(tools, position="sidebar", expanded=True)

# ──────────────────────────────────────────────
# Home page
# ──────────────────────────────────────────────
if pg.title == "Home":

    st.markdown("""
    <div class="hero-header">
        <h1>Operations Automation Dashboard</h1>
        <p>Internal tooling for order fulfillment, vendor portals, and business intelligence reporting.</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Stats ──
    st.markdown("""
    <div class="stats-row">
        <div class="stat-pill">
            <div class="stat-val">18</div>
            <div class="stat-lbl">Tools</div>
        </div>
        <div class="stat-pill">
            <div class="stat-val">4</div>
            <div class="stat-lbl">Platforms</div>
        </div>
        <div class="stat-pill">
            <div class="stat-val">5–10h</div>
            <div class="stat-lbl">Saved / week</div>
        </div>
        <div class="stat-pill">
            <div class="stat-val">2019</div>
            <div class="stat-lbl">In production</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.info("Select a tool from the sidebar to get started.", icon="ℹ️")

    # ── Amazon section ──
    st.markdown('<div class="section-label">Amazon Vendor Central</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="tool-grid">
        <div class="tool-card">
            <div class="card-icon">📦</div>
            <div class="card-title">Backorder Pipeline</div>
            <div class="card-desc">Pulls backorder data via Selenium, processes inventory levels, routes items to transfer/substitute/reject decisions, and generates a formatted HTML email reply.</div>
            <span class="card-badge badge-live">Live</span>
        </div>
        <div class="tool-card">
            <div class="card-icon">📄</div>
            <div class="card-title">Order Paperwork Generator</div>
            <div class="card-desc">Exports pack worksheets from NetSuite, validates item cross-references, splits per-PO, generates pack sheets and carton labels, then drafts an Outlook email with attachments.</div>
            <span class="card-badge badge-beta">Beta</span>
        </div>
        <div class="tool-card">
            <div class="card-icon">📈</div>
            <div class="card-title">Weekly Sales Report</div>
            <div class="card-desc">Pulls WTD/YTD sales data from Amazon Vendor Central, merges with NetSuite COGS data, generates charts, and assembles a formatted email report for management.</div>
            <span class="card-badge badge-beta">Beta</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Walmart section ──
    st.markdown('<div class="section-label">Walmart Vendor Portal</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="tool-grid">
        <div class="tool-card">
            <div class="card-icon">🛒</div>
            <div class="card-title">Order Confirmation Pipeline</div>
            <div class="card-desc">Pre-confirms transfer availability, confirms orders in the Walmart portal, and generates routed order uploads — with warehouse priority logic built in.</div>
            <span class="card-badge badge-live">Live</span>
        </div>
        <div class="tool-card">
            <div class="card-icon">🏷️</div>
            <div class="card-title">Pack Slip & Carton Labels</div>
            <div class="card-desc">Generates formatted pack slip and carton label files ready for warehouse use, including LTL/TL load documents and pallet labels.</div>
            <span class="card-badge badge-live">Live</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Analytics section ──
    st.markdown('<div class="section-label">Analytics & Reporting</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="tool-grid">
        <div class="tool-card">
            <div class="card-icon">📊</div>
            <div class="card-title">Inventory Health Report</div>
            <div class="card-desc">Pulls live inventory data from NetSuite via SuiteQL, calculates weeks-of-stock, flags at-risk items, and exports a formatted report for all accounts.</div>
            <span class="card-badge badge-stable">Stable</span>
        </div>
        <div class="tool-card">
            <div class="card-icon">🔧</div>
            <div class="card-title">Snap-on Pack Slip Generator</div>
            <div class="card-desc">Generates properly formatted pack slips for Snap-on SupplyWeb orders, matching their specific document requirements.</div>
            <span class="card-badge badge-stable">Stable</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Footer ──
    st.markdown("""
    <div class="dash-footer">
        Built with Streamlit · Python · NetSuite SuiteQL · Selenium
    </div>
    """, unsafe_allow_html=True)

else:
    pg.run()
