import streamlit as st
import logic
import pandas as pd
import plotly.graph_objects as go
import os

if st.session_state.get("role") != "admin":
    st.error("Access Denied: Admins Only")
    st.stop()

st.markdown('''
<style>
.main-header { background-color: #ffcad4; padding: 20px; text-align: center; border-radius: 10px; font-size: 28px; font-weight: 700; margin-bottom: 2rem;}
.metric-box { background: white; padding: 20px; border-radius: 12px; text-align: center; box-shadow: 0px 4px 12px rgba(0,0,0,0.08); border: 1px solid #ffe4e8;}
.metric-title { font-size:16px; color:#555; }
.metric-value { font-size:32px; font-weight:700; color:#ff6b81; }
[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-header">Admin Dashboard</div>', unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["Upload Timetable", "Utility Analysis", "Booking Logs"])

with tab1:
    st.subheader("Upload Base Timetables")
    st.info("Upload the weekly template. The system will automatically generate the full year.")
    
    upload_type = st.radio("Timetable Type:", ["Classroom", "Lab"], horizontal=True)
    uploaded_file = st.file_uploader("Upload .xlsx", type=["xlsx"])
    year = st.number_input("Year", min_value=2024, max_value=2050, value=2026)
    
    if st.button("Process & Generate Full Year"):
        if uploaded_file:
            with st.spinner("Processing file... This may take a few moments."):
                # 1. Save uploaded file temporarily
                file_location = os.path.join(logic.DATA_DIR, "raw_upload.xlsx")
                with open(file_location, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # 2. Call logic directly! No API delays.
                if upload_type == "Classroom":
                    success, msg = logic.generate_classroom_full_year(file_location, "FEST_Room Occupancy", year)
                else:
                    merged_file_path = logic.generate_vertically_merged_lab(file_location)
                    success, msg = logic.generate_lab_full_year(merged_file_path, year)
                    
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
        else:
            st.error("Please upload a file first.")

with tab2:
    st.subheader("Utility Analysis")
    ut_type = st.radio("Resource Type:", ["Classroom", "Lab"], horizontal=True, key="ut_radio")
    
    # Call logic directly! Instantly fetches resources.
    available_resources = logic.get_resource_names(ut_type)
    options = ["All"] + available_resources
    ut_resource = st.selectbox("Select Resource to Analyze:", options=options)
    
    if st.button("Generate Analysis"):
        with st.spinner("Calculating..."):
            # Call logic directly!
            data = logic.calculate_utility(ut_type, ut_resource)
            
            if "error" in data:
                st.error(data["error"])
            else:
                occupied = data["occupied"]
                free = data["free"]
                total = data["total"]
                
                if total == 0:
                    st.warning("No data found for this resource. Is the timetable uploaded?")
                else:
                    col1, col2 = st.columns([1,1.2])
                    with col1:
                        st.markdown(f'''
                        <div class="metric-box">
                        <div class="metric-title">Occupied Slots</div>
                        <div class="metric-value">{occupied}</div>
                        </div><br>
                        <div class="metric-box">
                        <div class="metric-title">Free Slots</div>
                        <div class="metric-value">{free}</div>
                        </div><br>
                        <div class="metric-box">
                        <div class="metric-title">Total Slots</div>
                        <div class="metric-value">{total}</div>
                        </div>
                        ''', unsafe_allow_html=True)

                    with col2:
                        fig = go.Figure(data=[go.Pie(
                            labels=["Occupied","Free"],
                            values=[occupied,free],
                            hole=0,
                            textinfo='none',
                            marker=dict(colors=["#EA5F89","#9B3192"]),
                            hovertemplate="<b>%{label}</b><br>Slots: %{value}<br>%{percent}<extra></extra>"
                        )])
                        fig.update_layout(
                            height=350,
                            showlegend=True,
                            legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5),
                            margin=dict(t=10,b=10,l=10,r=10)
                        )
                        st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("All Bookings")

    if st.button("Refresh Logs"):
        st.rerun()

    # Read the bookings file directly! No API delays.
    if os.path.exists(logic.BOOKINGS_FILE) and os.path.getsize(logic.BOOKINGS_FILE) > 0:
        df = pd.read_excel(logic.BOOKINGS_FILE)
        
        if not df.empty:
            header = st.columns([2,1,1,1,2,1,1])
            header[0].markdown("**Email**")
            header[1].markdown("**Type**")
            header[2].markdown("**Resource**")
            header[3].markdown("**Date**")
            header[4].markdown("**Time Slot**")
            header[5].markdown("**Purpose**")
            header[6].markdown("**Action**")

            st.markdown("---")

            for i, row in df.iterrows():
                cols = st.columns([2,1,1,1,2,1,1])
                cols[0].write(row["Email"])
                cols[1].write(row["Type"])
                cols[2].write(row["Resource"])
                cols[3].write(str(row["Date"]))
                cols[4].write(row["Time Slot"])
                cols[5].write(row["Purpose"])

                if cols[6].button("❌ Cancel", key=f"cancel_{i}"):
                    time_slots_list = str(row["Time Slot"]).split(" | ")
                    # Call logic directly to cancel!
                    cancel_res = logic.cancel_booking(
                        row["Email"], row["Type"], row["Resource"], str(row["Date"]), time_slots_list
                    )

                    if "error" not in cancel_res:
                        st.success("Booking cancelled successfully ✅")
                        st.rerun()
                    else:
                        st.error(cancel_res["error"])
                st.markdown("---")
        else:
            st.info("No bookings found yet.")
    else:
        st.info("No bookings found yet.")
