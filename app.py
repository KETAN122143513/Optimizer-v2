import streamlit as st
import pandas as pd
import pulp
import io

st.set_page_config(page_title="Airline Network Optimizer", layout="centered")
st.title("✈️ Airline Cargo Network Optimizer (Excel Upload)")

uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0).replace('-', 0).fillna(0)
        st.success("✅ Excel file loaded successfully!")

        all_od_paths = {}
        leg_capacities = {}

        for _, row in df.iterrows():
            od = row['Sector']
            cm = float(row['CM'])
            cargo_type = row['Cargo Type']
            market_size = float(row['Market Size'])
            cap = float(row['Capacity'])

            # AI Share logic
            if cargo_type == 'Direct':
                max_allocable = 0.5 * market_size
            else:
                max_allocable = 0.1 * market_size

            leg1 = row['Leg 1'] if pd.notna(row['Leg 1']) else None
            leg2 = row['Leg 2'] if pd.notna(row['Leg 2']) else None

            legs = []
            if leg1 and leg2:
                legs = [leg1, leg2]
            elif leg1:
                legs = [leg1]
            else:
                legs = [od]

            all_od_paths[od] = {
                'legs': legs,
                'cm': cm,
                'max_allocable': min(max_allocable, cap)
            }

            for leg in legs:
                leg_capacities[leg] = min(leg_capacities.get(leg, float('inf')), cap)

        # Optimization
        prob = pulp.LpProblem("NetworkCargoProfitMaximization", pulp.LpMaximize)
        x_od = pulp.LpVariable.dicts("CargoTons", all_od_paths.keys(), lowBound=0, cat='Continuous')
        prob += pulp.lpSum([x_od[od] * props['CM'] for od, props in all_od_paths.items()]), "TotalProfit"

        for leg, cap in leg_capacities.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs']]) <= cap

        for od, props in all_od_paths.items():
            prob += x_od[od] <= props['max_allocable']

        prob.solve()

        od_summary = []
        for v in prob.variables():
            if v.varValue > 0:
                od = v.name.replace("CargoTons_", "").replace("_", "-")
                tons = v.varValue
                cm = all_od_paths[od]['cm']
                profit = tons * cm
                od_summary.append({
                    'OD Pair': od,
                    'Cargo Tonnage': round(tons, 2),
                    'CM (₹/ton)': cm,
                    'Total Profit (₹)': round(profit, 2)
                })

        df_od_summary = pd.DataFrame(od_summary)

        records = []
        for od_row in df_od_summary.itertuples():
            od, tons, cm = od_row._1, od_row._2, od_row._3
            for leg in all_od_paths[od]['legs']:
                records.append({
                    'Flight Leg': leg,
                    'OD Contributor': od,
                    'OD CM (₹/ton)': cm,
                    'Cargo Tonnage': tons,
                    'Revenue from Leg (₹)': tons * cm
                })

        df_leg_detail = pd.DataFrame(records)
        df_leg_detail['Priority Type'] = "Fills Remaining"
        df_leg_detail['Fill Priority Rank'] = None

        for leg, group in df_leg_detail.groupby("Flight Leg"):
            sorted_group = group.sort_values(by=["Cargo Tonnage"], ascending=False)
            for rank, idx in enumerate(sorted_group.index, start=1):
                df_leg_detail.loc[idx, 'Fill Priority Rank'] = rank
                if len(group) == 1:
                    df_leg_detail.loc[idx, 'Priority Type'] = "Only OD"
                else:
                    df_leg_detail.loc[idx, 'Priority Type'] = "Based on Cargo Tonnage"

        df_leg_summary = df_leg_detail.groupby('Flight Leg')['Cargo Tonnage'].sum().reset_index(name='Total Tonnage (Tons)')
        total_profit = pulp.value(prob.objective)
        df_profit_note = pd.DataFrame([{
            'Flight Leg': 'TOTAL NETWORK PROFIT',
            'OD Contributor': '',
            'OD CM (₹/ton)': '',
            'Cargo Tonnage': '',
            'Revenue from Leg (₹)': round(total_profit, 2),
            'Priority Type': '',
            'Fill Priority Rank': ''
        }])

        tab1, tab2, tab3, tab4 = st.tabs(["📂 Input Sheet", "📦 OD Allocation", "✈️ Leg Breakdown", "📊 Summary & Download"])
        with tab1:
            st.dataframe(df)
        with tab2:
            st.dataframe(df_od_summary)
        with tab3:
            st.dataframe(df_leg_detail)
        with tab4:
            st.dataframe(df_leg_summary)
            st.markdown(f"### ✅ Total Network Profit: ₹ {round(total_profit, 2):,.2f}")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Raw_Input")
            df_od_summary.to_excel(writer, index=False, sheet_name="OD_Allocations")
            df_leg_detail.to_excel(writer, index=False, sheet_name="Leg_Breakdown")
            df_leg_summary.to_excel(writer, index=False, sheet_name="Leg_Summary")
            df_profit_note.to_excel(writer, index=False, sheet_name="Profit_Summary")
        output.seek(0)

        st.download_button("📥 Download Excel Report", data=output, file_name="Updated_Airline_Cargo_Report.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
