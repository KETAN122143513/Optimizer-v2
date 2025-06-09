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
        leg_total_capacities = {} # Renamed to be explicit about total leg capacity
        leg_tp_caps = {} # TP Cap for specific OD-Leg combinations
        leg_tp_master_caps = {} # TP Master Cap for indirect traffic on a leg

        # Initialize leg_total_capacities from direct routes
        for _, row in direct_routes.iterrows():
            od = None
            if 'O-D' in row and pd.notna(row['O-D']):
                od = str(row['O-D']).strip()
            elif 'Sector' in row and pd.notna(row['Sector']):
                od = str(row['Sector']).strip()
            if od:
                ai_cap = float(row['AI Cap'])
                leg_total_capacities[od] = ai_cap
                # Also add direct route to all_od_paths
                cm = float(row['CM'])
                ai_share = float(row['AI Share'])
                all_od_paths[od] = {
                    'legs': [od],
                    'cm': cm,
                    'max_allocable': min(ai_share, ai_cap),
                    'type': 'Direct'
                }


        # Process indirect routes
        for _, row in indirect_routes.iterrows():
            try:
                od = None
                if 'O-D' in row and pd.notna(row['O-D']):
                    od = str(row['O-D']).strip()
                elif 'Sector' in row and pd.notna(row['Sector']):
                    od = str(row['Sector']).strip()
                if not od:
                    continue
                cm = float(row['CM'])
                ai_share = float(row['AI Share'])
                tp_od_cap = ai_share
                max_allocable = min(ai_share, tp_od_cap) # This is a direct constraint on the OD itself
                
                legs = []
                leg1 = str(row['1st Leg O-D']).strip() if pd.notna(row['1st Leg O-D']) else None
                leg2 = str(row['2nd Leg O-D']).strip() if pd.notna(row['2nd Leg O-D']) else None
                legs = [leg1, leg2]

                all_od_paths[od] = {
                    'legs': [leg for leg in legs if leg is not None], # Store only valid legs
                    'cm': cm,
                    'max_allocable': max_allocable,
                    'type': 'Indirect'
                }

                for i, leg in enumerate(legs):
                    if leg is None:
                        continue
                    tp_cap = float(row[f'TP Cap {i+1}'])
                    master_cap = float(row[f'TP Master Cap {i+1}'])
                    # leg_cap = float(row[f'{i+1}st Leg AI Share']) # This was the problematic line for overall leg capacity

                    # Add to leg_tp_caps for OD-specific leg capacity
                    leg_tp_caps.setdefault(leg, []).append((od, tp_cap))
                    
                    # Update TP Master Cap for the leg
                    leg_tp_master_caps[leg] = master_cap

                    # Ensure the leg itself exists in leg_total_capacities,
                    # even if it's only used by indirect routes and not a direct route itself
                    # If direct routes define the primary capacity, these entries will be overridden.
                    # If not, you might need to infer or explicitly define these.
                    # For now, let's assume direct_routes sheet holds the definitive AI Cap for legs.
                    # If an indirect route uses a leg not in direct_routes, its total capacity would be master_cap
                    if leg not in leg_total_capacities:
                        st.warning(f"Leg '{leg}' used by indirect route '{od}' but not defined in Direct Routes. Assuming its total capacity will be limited by TP Master Cap: {master_cap}")
                        leg_total_capacities[leg] = master_cap # Fallback if not defined directly


            except Exception as e:
                st.error(f"Error processing indirect route row for OD {row.get('O-D', 'N/A')}: {e}")
                continue


        # Optimization
        prob = pulp.LpProblem("NetworkCargoProfitMaximization", pulp.LpMaximize)
        x_od = pulp.LpVariable.dicts("CargoTons", all_od_paths.keys(), lowBound=0, cat='Continuous')
        prob += pulp.lpSum([x_od[od] * props['cm'] for od, props in all_od_paths.items()]), "TotalProfit"

        # Leg capacity (Total traffic on a leg)
        for leg, cap in leg_total_capacities.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs']]) <= cap, f"Leg_Total_Capacity_{leg}"

        # Max allocable per OD (Market Size / AI Share / TP Cap OD for the OD itself)
        for od, props in all_od_paths.items():
            prob += x_od[od] <= props['max_allocable'], f"Max_Allocable_{od}"

        # TP OD-Leg cap (Specific OD's usage of a leg)
        for leg, tp_list in leg_tp_caps.items():
            for od, cap in tp_list:
                prob += x_od[od] <= cap, f"TP_OD_Leg_Cap_{od}_{leg}"

        # TP Master cap (Total indirect traffic on a leg)
        for leg, master_cap in leg_tp_master_caps.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs'] and props['type'] == 'Indirect']) <= master_cap, f"TP_Master_Cap_{leg}"

        prob.solve()

        od_summary = []
        for v in prob.variables():
            if v.varValue > 0:
                od = v.name.replace("CargoTons_", "").replace("_", "-")
                # Need to handle special characters in OD names if they exist and are not just hyphens
                # For example, if 'DEL-JFK' becomes 'DEL_JFK' in pulp, then 'DEL-JFK' should be the key in all_od_paths
                # The .replace("_", "-") here assumes all underscores in the var name were originally hyphens
                # It's safer to map back to the original keys if possible.
                original_od_key = next((k for k in all_od_paths.keys() if k == od), None)
                if original_od_key is None: # If direct match fails, try replacing hyphens
                    original_od_key = next((k for k in all_od_paths.keys() if k.replace('-', '_') == od.replace('-', '_')), None)


                if original_od_key in all_od_paths:
                    tons = v.varValue
                    cm = all_od_paths[original_od_key]['cm']
                    profit = tons * cm
                    od_summary.append({
                        'OD Pair': original_od_key,
                        'Cargo Tonnage': round(tons, 2),
                        'CM (₹/ton)': cm,
                        'Total Profit (₹)': round(profit, 2)
                    })
                else:
                    st.warning(f"Could not find original OD key for variable: {v.name}. Skipping.")


        df_od_summary = pd.DataFrame(od_summary)

        records = []
        for od_row in df_od_summary.itertuples():
            od, tons, cm = od_row._1, od_row._2, od_row._3
            # Use original_od_key to fetch legs from all_od_paths
            original_od_key = next((k for k in all_od_paths.keys() if k == od), None)
            if original_od_key:
                for leg in all_od_paths[original_od_key]['legs']:
                    records.append({
                        'Flight Leg': leg,
                        'OD Contributor': od,
                        'OD CM (₹/ton)': cm,
                        'Cargo Tonnage': tons,
                        'Revenue from Leg (₹)': tons * cm
                    })
            else:
                st.warning(f"Could not find original OD key '{od}' for leg breakdown. Skipping.")


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
