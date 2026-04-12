import streamlit as st
import requests
import pandas as pd
import plotly.graph_objects as go

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
                files = {"file": (uploaded_file.name, uploaded_file.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                data = {"year": year}
                
                if upload_type == "Classroom":
                    # For classroom, we also send the default sheet name
                    data["sheet_name"] = "FEST_Room Occupancy" 
                    url = "https://smart-class-api-xez6.onrender.com/admin/upload/classroom"
                else:
                    url = "https://smart-class-api-xez6.onrender.com/admin/upload/lab"
                    
                res = requests.post(url, files=files, data=data)
                
                if res.status_code == 200:
                    st.success(res.json().get("message", "Processed successfully!"))
                else:
                    st.error(res.json().get("detail", "An error occurred during processing."))
        else:
            st.error("Please upload a file first.")

with tab2:
    st.subheader("Utility Analysis")
    ut_type = st.radio("Resource Type:", ["Classroom", "Lab"], horizontal=True, key="ut_radio")
    
    # Fetch resource names dynamically from the backend
    try:
        res_req = requests.get(f"https://smart-class-api-xez6.onrender.com/resources?resource_type={ut_type}")
        available_resources = res_req.json().get("resources", []) if res_req.status_code == 200 else []
    except requests.exceptions.ConnectionError:
        available_resources = []
        st.error("Could not connect to backend to fetch resources.")
    
    # Create the proper dropdown menu
    options = ["All"] + available_resources
    ut_resource = st.selectbox("Select Resource to Analyze:", options=options)
    
    if st.button("Generate Analysis"):
        with st.spinner("Calculating..."):
            payload = {"resource_type": ut_type, "selected_resource": ut_resource}
            res = requests.post("https://smart-class-api-xez6.onrender.com/utility", json=payload)
            
            if res.status_code == 200:
                data = res.json()
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
            else:
                st.error(res.json().get("detail", "Error calculating utility."))

with tab3:
    st.subheader("All Bookings")

    if st.button("Refresh Logs"):
        res = requests.get("https://smart-class-api-xez6.onrender.com/admin/bookings")

        if res.status_code == 200:
            logs = res.json()

            if logs:
                df = pd.DataFrame(logs)

                st.dataframe(df, use_container_width=True, hide_index=True)

                st.markdown("### ❌ Cancel Booking")

                booking_index = st.number_input(
                    "Enter booking index to cancel",
                    min_value=0,
                    max_value=len(df)-1,
                    step=1
                )

                if st.button("Cancel Booking"):
                    payload = {"index": int(booking_index)}

                    cancel_res = requests.post(
                        "https://smart-class-api-xez6.onrender.com/admin/cancel",
                        json=payload
                    )

                    if cancel_res.status_code == 200:
                        st.success("Booking cancelled successfully ✅")
                    else:
                        st.error(cancel_res.json().get("detail", "Error cancelling booking"))

            else:
                st.info("No bookings found yet.")
        else:
            st.error("Could not fetch bookings.")
