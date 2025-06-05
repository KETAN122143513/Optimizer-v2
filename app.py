import streamlit as st
import pandas as pd
import pulp
import io

st.set_page_config(page_title="Airline Cargo Network Optimizer", layout="centered")
st.title("✈️ Airline Cargo Network Optimizer (Excel Upload)")

uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type="xlsx")

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        direct_routes = xls.parse(sheet_name=0).replace("-", 0).fillna(0)
        indirect_routes = xls.parse(sheet_name=1).replace("-", 0).fillna(0)
        st.success("✅ Excel file loaded successfully!")

        all_od_paths = {}
        leg_capacities = {}
        leg_tp_caps = {}  # max TP cargo for any OD on a leg
        leg_tp_master_caps = {}  # total TP cargo across all ODs per leg
        od_leg_tp_caps = {}  # individual OD-leg constraint: min(10% leg cap, 12% AI Share)

        # Process direct routes
        for _, row in direct_routes.iterrows():
            try:
                od = row['O-D']
                cm = float(row['CM'])
                ai_share = float(row['AI Share']) if pd.notna(row['AI Share']) else 0
                capacity = float(row['AI Cap']) if pd.notna(row['AI Cap']) else 0
                tp_cap = float(row['TP Cap']) if pd.notna(row['TP Cap']) else 0
                tp_master_cap = float(row['TP Master Cap']) if pd.notna(row['TP Master Cap']) else 0

                all_od_paths[od] = {
                    'legs': [od],
                    'cm': cm,
                    'max_allocable': min(ai_share, capacity),
                    'type': 'Direct'
                }
                leg_capacities[od] = capacity
                leg_tp_caps[od] = tp_cap
                leg_tp_master_caps[od] = tp_master_cap
            except:
                continue

        # Process indirect routes
        for _, row in indirect_routes.iterrows():
            try:
                od = row['O-D']
                cm = float(row['CM'])
                ai_share = float(row['AI Share']) if pd.notna(row['AI Share']) else 0
                max_allocable = ai_share
                legs = [row['1st Leg O-D'], row['2nd Leg O-D']]
                all_od_paths[od] = {
                    'legs': legs,
                    'cm': cm,
                    'max_allocable': max_allocable,
                    'type': 'Indirect'
                }

                for i, leg in enumerate(legs):
                    cap_col = f'{i+1}st Leg AI Cap'
                    leg_cap = float(row[cap_col]) if pd.notna(row[cap_col]) else 0
                    leg_capacities[leg] = min(leg_capacities.get(leg, float('inf')), leg_cap)
                    leg_tp_caps[leg] = 0.1 * leg_cap
                    leg_tp_master_caps[leg] = 0.3 * leg_cap
                    od_leg_tp_caps[(od, leg)] = min(0.1 * leg_cap, 0.12 * ai_share)
            except:
                continue

        # Optimization
        prob = pulp.LpProblem("NetworkCargoProfitMaximization", pulp.LpMaximize)
        x_od = pulp.LpVariable.dicts("CargoTons", all_od_paths.keys(), lowBound=0, cat='Continuous')
        prob += pulp.lpSum([x_od[od] * props['cm'] for od, props in all_od_paths.items()]), "TotalProfit"

        # Flight leg capacity constraint
        for leg, cap in leg_capacities.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs']]) <= cap

        # OD max allocable constraint
        for od, props in all_od_paths.items():
            prob += x_od[od] <= props['max_allocable']

        # TP constraints
        for leg in leg_tp_caps:
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs'] and props['type'] == 'Indirect']) <= leg_tp_master_caps[leg]

        for (od, leg), cap in od_leg_tp_caps.items():
            prob += x_od[od] <= cap

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
            sorted_group = group.copy()
            sorted_group = sorted_group.sort_values(by=["Cargo Tonnage"], ascending=False)
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

        tab1, tab2, tab3, tab4 = st.tabs(["📂 Input Sheets", "📦 OD Allocation", "✈️ Leg Breakdown", "📊 Summary & Download"])
        with tab1:
            st.subheader("Direct Routes (Input)")
            st.dataframe(direct_routes)
            st.subheader("Indirect Routes (Input)")
            st.dataframe(indirect_routes)
        with tab2:
            st.dataframe(df_od_summary)
        with tab3:
            st.dataframe(df_leg_detail)
        with tab4:
            st.dataframe(df_leg_summary)
            st.markdown(f"### ✅ Total Network Profit: ₹ {round(total_profit, 2):,.2f}")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            direct_routes.to_excel(writer, index=False, sheet_name="Direct_Routes_Input")
            indirect_routes.to_excel(writer, index=False, sheet_name="Indirect_Routes_Input")
            df_od_summary.to_excel(writer, index=False, sheet_name="OD_Allocations")
            df_leg_detail.to_excel(writer, index=False, sheet_name="Leg_Breakdown")
            df_leg_summary.to_excel(writer, index=False, sheet_name="Leg_Summary")
            df_profit_note.to_excel(writer, index=False, sheet_name="Profit_Summary")
        output.seek(0)

        st.download_button("📥 Download Excel Report", data=output, file_name="Airline_Cargo_Report.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
