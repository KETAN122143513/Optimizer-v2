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
        leg_capacities = {} # Overall capacity for a flight leg (AI Cap for direct, AI Share for indirect leg)
        leg_tp_master_caps = {} # TP Master Cap for indirect legs

        # Process direct routes
        for _, row in direct_routes.iterrows():
            try:
                od = row.get('O-D') or row.get('Sector')
                if not od:
                    continue
                cm = float(row['CM'])
                ai_share = float(row['AI Share'])
                ai_cap = float(row['AI Cap'])
                market_size = float(row['Market Size'])

                all_od_paths[od] = {
                    'legs': [od],
                    'cm': cm,
                    'ai_share': ai_share,
                    'market_size': market_size,
                    'type': 'Direct'
                }
                # For direct routes, the leg's capacity is its AI Cap
                leg_capacities[od] = ai_cap
            except Exception as e:
                st.warning(f"Skipping direct route row due to error: {e} in row: {row.to_dict()}")
                continue

        # Process indirect routes
        for _, row in indirect_routes.iterrows():
            try:
                od = row.get('O-D') or row.get('Sector')
                if not od:
                    continue
                cm = float(row['CM'])
                ai_share = float(row['AI Share']) # This is the AI Share for the overall O-D
                market_size = float(row['Market Size'])
                
                leg1_od = row['1st Leg O-D']
                leg1_ai_share = float(row['1st Leg AI Share'])
                leg1_ai_cap = float(row['1st Leg AI Cap']) # This is not used as a strict cap on the leg, rather the AI share of the leg
                leg1_cm = float(row['1st Leg CM'])
                leg1_tp_cap = float(row['TP Cap 1'])
                leg1_tp_master_cap = float(row['TP Master Cap 1'])

                leg2_od = row['2nd Leg O-D']
                leg2_ai_share = float(row['2nd Leg AI Share'])
                leg2_ai_cap = float(row['2nd Leg AI Cap']) # This is not used as a strict cap on the leg, rather the AI share of the leg
                leg2_cm = float(row['2nd Leg CM'])
                leg2_tp_cap = float(row['TP Cap 2'])
                leg2_tp_master_cap = float(row['TP Master Cap 2'])

                legs = [leg1_od, leg2_od]

                # Determine the true max allocable for the O-D considering all caps
                max_allocable_od = min(ai_share, leg1_tp_cap, leg2_tp_cap)

                all_od_paths[od] = {
                    'legs': legs,
                    'cm': cm,
                    'ai_share': ai_share,
                    'market_size': market_size,
                    'type': 'Indirect',
                    'leg1_tp_cap': leg1_tp_cap,
                    'leg2_tp_cap': leg2_tp_cap,
                    'max_allocable': max_allocable_od # This is the overall max allocable for the O-D
                }

                # Update leg capacities for individual legs. For indirect, we use the AI Share of the leg.
                # If a direct route shares a leg with an indirect, the AI Cap of the direct takes precedence for the leg.
                # Here we are concerned with the capacity of the *leg itself*
                leg_capacities[leg1_od] = min(leg_capacities.get(leg1_od, float('inf')), leg1_ai_share)
                leg_capacities[leg2_od] = min(leg_capacities.get(leg2_od, float('inf')), leg2_ai_share)

                # Store TP Master Caps for later use in constraints
                leg_tp_master_caps[leg1_od] = leg1_tp_master_cap
                leg_tp_master_caps[leg2_od] = leg2_tp_master_cap

            except Exception as e:
                st.warning(f"Skipping indirect route row due to error: {e} in row: {row.to_dict()}")
                continue
        
        # Merge AI Cap for direct legs into leg_capacities if it's higher or unique
        for od, props in all_od_paths.items():
            if props['type'] == 'Direct':
                leg_capacities[od] = min(leg_capacities.get(od, float('inf')), props.get('ai_cap', float('inf'))) # Ensuring direct AI Cap is considered


        # Optimization
        prob = pulp.LpProblem("NetworkCargoProfitMaximization", pulp.LpMaximize)
        x_od = pulp.LpVariable.dicts("CargoTons", all_od_paths.keys(), lowBound=0, cat='Continuous')

        # Objective function: Maximize total profit
        prob += pulp.lpSum([x_od[od] * props['cm'] for od, props in all_od_paths.items()]), "TotalProfit"

        # Constraints

        # 1. Overall leg capacity constraint (AI Cap for direct routes, AI Share for indirect legs)
        for leg, cap in leg_capacities.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs']]) <= cap, f"Leg_Capacity_{leg}"

        # 2. Max allocable per OD pair (based on AI Share for direct, and min of AI Share and both TP Caps for indirect)
        for od, props in all_od_paths.items():
            if props['type'] == 'Direct':
                prob += x_od[od] <= props['ai_share'], f"OD_Max_Allocable_{od}"
            else: # Indirect routes
                prob += x_od[od] <= props['max_allocable'], f"OD_Max_Allocable_{od}"


        # 3. TP Master Cap constraint for indirect legs
        # This constraint states that the sum of all indirect cargo using a particular leg
        # must not exceed the TP Master Cap for that leg.
        for leg, master_cap in leg_tp_master_caps.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs'] and props['type'] == 'Indirect']) <= master_cap, f"TP_Master_Cap_{leg}"
        
        # Solve the problem
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
            # Sort by CM first, then Cargo Tonnage if CM is same (though CM is already captured by profit max)
            sorted_group = group.sort_values(by=["Cargo Tonnage"], ascending=False)
            for rank, idx in enumerate(sorted_group.index, start=1):
                df_leg_detail.loc[idx, 'Fill Priority Rank'] = rank
                if len(group) == 1:
                    df_leg_detail.loc[idx, 'Priority Type'] = "Only OD"
                else:
                    df_leg_detail.loc[idx, 'Priority Type'] = "Based on Cargo Tonnage" # The solver optimizes for CM already

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
            st.subheader("Optimal OD Allocations")
            st.dataframe(df_od_summary)
        with tab3:
            st.subheader("Cargo Breakdown per Flight Leg")
            st.dataframe(df_leg_detail)
        with tab4:
            st.subheader("Total Tonnage per Flight Leg")
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
