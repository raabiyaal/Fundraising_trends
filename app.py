import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from dash import Dash, Input, Output, dcc, html

DATA_PATH = Path(__file__).with_name("Fundraising Data.xlsx")


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

    # Helper to coerce numeric values
    def to_number(series: pd.Series) -> pd.Series:
        return pd.to_numeric(
            series.astype(str).str.replace(r"[$,]", "", regex=True),
            errors="coerce",
        )

    header_map = {
        "Year": "Year",
        "Number of Funds": "Number of Funds",
        "Total Amount": "Amount Closed",
        "Average Fund Size": "Average Fund Size",
    }

    if all(k in df.columns for k in header_map.keys()):
        out = df[list(header_map.keys())].rename(columns=header_map)
    else:
        cols = list(df.columns)
        if len(cols) < 4:
            raise ValueError("Expected at least 4 columns in the data file.")
        out = pd.DataFrame(
            {
                "Year": df[cols[0]],
                "Number of Funds": df[cols[1]],
                "Amount Closed": df[cols[2]],
                "Average Fund Size": df[cols[3]],
            }
        )

    for c in ["Year", "Number of Funds", "Amount Closed", "Average Fund Size"]:
        out[c] = to_number(out[c])

    out = out.dropna(subset=["Year"])
    out["Year"] = out["Year"].astype(int)
    out = out.sort_values("Year")
    return out


data = load_data(DATA_PATH)

app = Dash(__name__)


def make_figure(metric: str) -> go.Figure:
    bar_color = "#0000FF"
    line_colors = {
        "Number of Funds": "#00A651",
        "Average Fund Size": "#A16CFF",
    }

    line_color = line_colors[metric]
    line_prefix = "$" if metric == "Average Fund Size" else ""

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
    fig.add_trace(
        go.Scatter(
            x=data["Year"],
            y=data[metric],
            name=line_name,
            mode="lines+markers",
            yaxis="y2",
            line=dict(color=line_color, width=3),
            marker=dict(color=line_color, size=7),
            hovertemplate=(
                f"Year %{{x}}<br>{metric}: {line_prefix}%{{y:,.0f}}"
                + ("M" if metric == "Average Fund Size" else "")
                + "<extra></extra>"
            ),
        )
    )

    right_title = (
        "Number of Funds Closed"
        if metric == "Number of Funds"
        else "Average Fund Size ($ millions)"
    )

    fig.update_layout(
        title="",
        template="simple_white",
        font=dict(family="Georgia", size=12),
        bargap=0.35,
        margin=dict(t=30, b=20, l=20, r=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            font=dict(size=12),
        ),
        xaxis=dict(
            title="",
            tickmode="linear",
            tick0=2006,
            dtick=1,
            range=[2005.5, 2025.5],
            tickangle=0,
            tickfont=dict(size=10),
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
    return fig


app.layout = html.Div(
    [
        dcc.Dropdown(
            id="line-set",
            options=[
                {"label": "Number of Funds", "value": "Number of Funds"},
                {"label": "Average Fund Size", "value": "Average Fund Size"},
            ],
            value="Number of Funds",
            clearable=False,
            style={"width": "50%", "margin": "20px auto"},
        ),
        dcc.Graph(
            id="fundraising-graph",
            style={"width": "90%", "height": "80vh", "margin": "auto"},
        ),
    ]
)


@app.callback(Output("fundraising-graph", "figure"), Input("line-set", "value"))
def update_graph(metric: str) -> go.Figure:
    return make_figure(metric)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
