import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(page_title="High-Yield Debt Fundraising", layout="wide")

DATA_PATH = Path(__file__).with_name("Fundraising Data.xlsx")

@st.cache_data
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Drop fully empty columns
    df = df.dropna(axis=1, how="all")

    # Drop columns that are just currency symbols or empty strings
    def is_symbol_col(series: pd.Series) -> bool:
        if series.dtype == object:
            vals = series.dropna().astype(str).str.strip().unique()
            return len(vals) == 1 and vals[0] in {"$", "USD", "US$", ""}
        return False

    df = df[[c for c in df.columns if not is_symbol_col(df[c])]]

    # If columns are unnamed and first row looks like headers, promote it
    if all(str(c).startswith("Unnamed") for c in df.columns):
        first_row = df.iloc[0].astype(str).str.strip()
        if first_row.str.contains(r"[A-Za-z]").any():
            df.columns = first_row
            df = df.iloc[1:]

    # Helper to coerce numeric values
    def to_number(series: pd.Series) -> pd.Series:
        return pd.to_numeric(
            series.astype(str).str.replace(r"[$,]", "", regex=True),
            errors="coerce",
        )

    cols = list(df.columns)
    year_col = None
    num_col = None
    amount_col = None
    avg_col = None

    # Prefer exact headers from your sheet
    header_map = {
        "Year": "Year",
        "Number of Funds": "Number of Funds",
        "Total Amount": "Amount Closed",
        "Average Fund Size": "Average Fund Size",
    }

    if all(k in df.columns for k in header_map.keys()):
        out = df[list(header_map.keys())].rename(columns=header_map)
    else:
        for c in cols:
            name = str(c).lower()
            if year_col is None and "year" in name:
                year_col = c
            elif num_col is None and ("number" in name or "funds" in name) and "avg" not in name:
                num_col = c
            elif amount_col is None and ("amount" in name or "closed" in name or "total" in name):
                amount_col = c
            elif avg_col is None and ("average" in name or "avg" in name):
                avg_col = c

        # Fallback to positional mapping if any are missing
        if year_col is None and len(cols) >= 1:
            year_col = cols[0]
        if num_col is None and len(cols) >= 2:
            num_col = cols[1]
        if amount_col is None and len(cols) >= 3:
            amount_col = cols[2]
        if avg_col is None and len(cols) >= 4:
            avg_col = cols[3]

        out = pd.DataFrame(
            {
                "Year": to_number(df[year_col]),
                "Number of Funds": to_number(df[num_col]) if num_col is not None else None,
                "Amount Closed": to_number(df[amount_col]) if amount_col is not None else None,
                "Average Fund Size": to_number(df[avg_col]) if avg_col is not None else None,
            }
        )

    # Coerce numeric values if we took the header path above
    for c in ["Year", "Number of Funds", "Amount Closed", "Average Fund Size"]:
        out[c] = to_number(out[c])

    out = out.dropna(subset=["Year"])
    out["Year"] = out["Year"].astype(int)
    out = out.sort_values("Year")
    return out


st.title("High-Yield Debt Funds Closed by Year")

if not DATA_PATH.exists():
    st.error(f"Data file not found: {DATA_PATH.name}")
    st.stop()

try:
    data = load_data(DATA_PATH)
except Exception as exc:
    st.error("Failed to read the Excel file. Please check the file format and headers.")
    st.exception(exc)
    st.stop()

metric = st.selectbox(
    "Secondary y-axis",
    ["Number of Funds", "Average Fund Size"],
    index=0,
)

bar_color = "#0000FF"
line_colors = {
    "Number of Funds": "#00A651",
    "Average Fund Size": "#A16CFF",
}

line_color = line_colors[metric]
secondary_values = data[metric]

fig = go.Figure()

fig.add_trace(
    go.Bar(
        x=data["Year"],
        y=data["Amount Closed"],
        name="Amount Closed ($ millions)",
        marker_color=bar_color,
        hovertemplate="Year %{x}<br>Amount Closed: $%{y:,.0f}M<extra></extra>",
    )
)

line_name = metric + (" ($ millions)" if metric == "Average Fund Size" else "")
line_prefix = "$" if metric == "Average Fund Size" else ""

fig.add_trace(
    go.Scatter(
        x=data["Year"],
        y=secondary_values,
        name=line_name,
        mode="lines+markers",
        yaxis="y2",
        line=dict(color=line_color, width=3),
        marker=dict(color=line_color, size=7),
        hovertemplate=f"Year %{{x}}<br>{metric}: {line_prefix}%{{y:,.0f}}" + ("M" if metric == "Average Fund Size" else "") + "<extra></extra>",
    )
)

right_title = "Number of Funds Closed" if metric == "Number of Funds" else "Average Fund Size ($ millions)"

fig.update_layout(
    title="High-Yield Debt Funds Closed by Year, for the Period 2006 through 2025",
    font=dict(family="Georgia", size=14),
    plot_bgcolor="white",
    bargap=0.35,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    margin=dict(l=60, r=60, t=80, b=60),
    xaxis=dict(
        title="",
        tickmode="linear",
        tick0=2006,
        dtick=1,
        range=[2005.5, 2025.5],
        tickfont=dict(size=12),
        showgrid=False,
    ),
    yaxis=dict(
        title=dict(text="Amount Closed ($ millions)", font=dict(color=bar_color)),
        tickprefix="$",
        separatethousands=True,
        tickformat=",d",
        tickfont=dict(color=bar_color),
        gridcolor="rgba(0,0,0,0.15)",
        zerolinecolor="rgba(0,0,0,0.2)",
    ),
    yaxis2=dict(
        title=dict(text=right_title, font=dict(color=line_color)),
        overlaying="y",
        side="right",
        tickprefix=line_prefix,
        separatethousands=True,
        tickformat=",d",
        tickfont=dict(color=line_color),
        showgrid=False,
    ),
)

st.plotly_chart(fig, use_container_width=False, width=900, height=800)

st.caption("Source: High-Yield Fund Database, Green Street and Instructor's calculations.")
