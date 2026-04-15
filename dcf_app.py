from io import BytesIO
import seaborn as sns
import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide", initial_sidebar_state="expanded")

st.title("📊 DCF Valuation Dashboard")

# -------------------------------
# SIDEBAR
# -------------------------------
st.sidebar.header("Global Inputs")

company_name = st.sidebar.text_input("Enter Company Name", "")
start_year = st.sidebar.number_input("Starting Year", value=2021)

net_debt = st.sidebar.number_input("Net Debt", value=0.0)
shares = st.sidebar.number_input("Shares Outstanding", value=1000.0)
market_price = st.sidebar.number_input("Market Price", value=0.0)

discount_rate = st.sidebar.slider("WACC (%)", 5.0, 25.0, 15.0) / 100
terminal_growth = st.sidebar.slider("Terminal Growth (%)", 0.0, 6.0, 4.0) / 100
tax_rate = st.sidebar.slider("Tax Rate (%)", 0.0, 50.0, 29.0) / 100
years = st.sidebar.slider("Forecast Years", 3, 10, 5)

# -------------------------------
# DCF FUNCTION
# -------------------------------
def dcf_model_full(
    revenue, growth_rate, ebit_margin, tax_rate,
    capex_pct, wc_pct, discount_rate, terminal_growth, years
):
    fcf_list = []
    revenues_proj = []
    prev_revenue = revenue

    for t in range(1, years + 1):
        revenue = revenue * (1 + growth_rate)
        delta_revenue = revenue - prev_revenue

        ebit = revenue * ebit_margin
        nopat = ebit * (1 - tax_rate)

        capex = revenue * capex_pct
        change_wc = delta_revenue * wc_pct

        fcf = nopat - capex - change_wc

        fcf_list.append(fcf)
        revenues_proj.append(revenue)

        prev_revenue = revenue

    discounted_fcf = [
        fcf / ((1 + discount_rate) ** t)
        for t, fcf in enumerate(fcf_list, start=1)
    ]

    terminal_value = (
        fcf_list[-1] * (1 + terminal_growth)
    ) / (discount_rate - terminal_growth)

    discounted_tv = terminal_value / ((1 + discount_rate) ** years)

    total_value = sum(discounted_fcf) + discounted_tv

    return total_value, fcf_list, revenues_proj


# -------------------------------
# TABS
# -------------------------------
tab1, tab2, tab3 = st.tabs(["📥 Data Input", "📊 Valuation", "📈 Sensitivity & Charts"])

# -------------------------------
# TAB 1 — DATA INPUT
# -------------------------------
with tab1:
    st.header("Enter Historical Data")

    years_hist = [int(start_year) + i for i in range(5)]

    revenues = []
    ebit_list = []
    wc_changes = []
    capex_list = []

    for i, year in enumerate(years_hist):
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            rev = st.number_input(f"Revenue {year}", key=f"rev_{i}", value=0.0)
        with col2:
            ebit = st.number_input(f"EBIT {year}", key=f"ebit_{i}", value=0.0)
        with col3:
            wc = st.number_input(f"WC Change {year}", key=f"wc_{i}", value=0.0)
        with col4:
            capex = st.number_input(f"Capex {year}", key=f"capex_{i}", value=0.0)

        revenues.append(rev)
        ebit_list.append(ebit)
        wc_changes.append(wc)
        capex_list.append(capex)


# -------------------------------
# ASSUMPTIONS
# -------------------------------
growth_rates = []
wc_percentages = []
ebit_margins = []
capex_pct_list = []

for i in range(1, len(revenues)):
    if revenues[i-1] != 0:
        growth_rates.append((revenues[i] - revenues[i-1]) / revenues[i-1])
        delta_rev = revenues[i] - revenues[i-1]
        if delta_rev != 0:
            wc_percentages.append(wc_changes[i] / delta_rev)

for i in range(len(revenues)):
    if revenues[i] != 0:
        ebit_margins.append(ebit_list[i] / revenues[i])
        capex_pct_list.append(capex_list[i] / revenues[i])

avg_growth = np.mean(growth_rates) if growth_rates else 0
avg_wc_pct = np.mean(wc_percentages) if wc_percentages else 0
avg_ebit_margin = np.mean(ebit_margins) if ebit_margins else 0
avg_capex_pct = np.mean(capex_pct_list) if capex_pct_list else 0

# -------------------------------
# TAB 2 — VALUATION
# -------------------------------
with tab2:
    st.header("Valuation")

    growth_rate = st.slider("Growth Rate (%)", 0.0, 20.0, float(avg_growth*100)) / 100
    ebit_margin = st.slider("EBIT Margin (%)", 0.0, 60.0, float(avg_ebit_margin*100)) / 100
    capex_pct = st.slider("Capex (%)", 0.0, 20.0, float(avg_capex_pct*100)) / 100
    wc_pct = st.slider("WC (%)", -10.0, 20.0, float(avg_wc_pct*100)) / 100

    revenue = revenues[-1] if revenues else 0

    if revenue > 0 and shares != 0:
        ev, fcf_list, revenues_proj = dcf_model_full(
            revenue, growth_rate, ebit_margin, tax_rate,
            capex_pct, wc_pct, discount_rate, terminal_growth, years
        )

        equity = ev - net_debt
        per_share = equity / shares
        mos = (per_share - market_price) / market_price if market_price != 0 else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("Fair Value", round(per_share, 2))
        col2.metric("Market Price", market_price)
        col3.metric("Margin of Safety", f"{round(mos*100,2)}%")

        if mos > 0.15:
            st.success("Undervalued")
        elif mos < -0.15:
            st.error("Overvalued")
        else:
            st.warning("Fairly Valued")

#-----------------------------------
# ------ Excel Export Code ---------
# ----------------------------------

def create_excel_download(company_name, per_share, market_price, mos,
                          growth_rate, ebit_margin, capex_pct, wc_pct,
                          discount_rate, terminal_growth, df):

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        # ---- Summary Sheet ----
        summary = pd.DataFrame({
            "Metric": [
                "Company",
                "Fair Value",
                "Market Price",
                "Margin of Safety",
                "Growth Rate",
                "EBIT Margin",
                "Capex %",
                "WC %",
                "WACC",
                "Terminal Growth"
            ],
            "Value": [
                company_name,
                round(per_share, 2),
                market_price,
                round(mos*100, 2),
                round(growth_rate*100, 2),
                round(ebit_margin*100, 2),
                round(capex_pct*100, 2),
                round(wc_pct*100, 2),
                round(discount_rate*100, 2),
                round(terminal_growth*100, 2)
            ]
        })

        summary.to_excel(writer, sheet_name="Summary", index=False)

        # ---- Sensitivity Sheet ----
        df.to_excel(writer, sheet_name="Sensitivity")

    return output.getvalue()

# -------------------------------
# TAB 3 — SENSITIVITY + CHARTS
# -------------------------------
with tab3:
    st.header("Sensitivity & Charts")

    revenue = revenues[-1] if revenues else 0

    if revenue > 0 and shares != 0:

        # -------- DCF CALCULATION --------
        ev, fcf_list, revenues_proj = dcf_model_full(
            revenue, growth_rate, ebit_margin, tax_rate,
            capex_pct, wc_pct, discount_rate, terminal_growth, years
        )

        equity = ev - net_debt
        per_share = equity / shares
        mos = (per_share - market_price) / market_price if market_price != 0 else 0

        # -------- SENSITIVITY TABLE --------
        growth_range = np.linspace(growth_rate - 0.02, growth_rate + 0.02, 5)
        wacc_range = np.linspace(discount_rate - 0.02, discount_rate + 0.02, 5)

        table = []

        for g in growth_range:
            row = []
            for w in wacc_range:
                ev_temp, _, _ = dcf_model_full(
                    revenue, g, ebit_margin, tax_rate,
                    capex_pct, wc_pct, w, terminal_growth, years
                )

                equity_temp = ev_temp - net_debt
                per_share_temp = equity_temp / shares

                row.append(round(per_share_temp, 1))

            table.append(row)

        df = pd.DataFrame(
            table,
            index=[round(g*100, 1) for g in growth_range],
            columns=[round(w*100, 1) for w in wacc_range]
        )

        # -------- TABLE + HEATMAP --------
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Sensitivity Table")
            st.dataframe(df)

        with col2:
            st.subheader("Heatmap")

            fig, ax = plt.subplots()

            sns.heatmap(
                df,
                annot=True,
                fmt=".1f",
                cmap="RdYlGn",
                ax=ax
            )

            ax.set_xlabel("WACC (%)")
            ax.set_ylabel("Growth (%)")

            st.pyplot(fig)

        # -------- FCF CHART --------
        st.subheader("Projected Free Cash Flow")

        fig2, ax2 = plt.subplots()
        ax2.plot(fcf_list, marker='o')
        ax2.set_title("FCF Projection")
        ax2.set_xlabel("Year")
        ax2.set_ylabel("FCF")

        st.pyplot(fig2)

        # -------- BASE CASE INFO --------
        st.info(
            f"Base Case → Growth: {round(growth_rate*100,1)}%, "
            f"WACC: {round(discount_rate*100,1)}%"
        )

        # -------- EXPORT TO EXCEL --------
        st.subheader("Download Results")

        excel_data = create_excel_download(
            company_name,
            per_share,
            market_price,
            mos,
            growth_rate,
            ebit_margin,
            capex_pct,
            wc_pct,
            discount_rate,
            terminal_growth,
            df
        )

        st.download_button(
            label="Download Excel Report",
            data=excel_data,
            file_name=f"{company_name}_DCF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please enter valid revenue and shares data.")