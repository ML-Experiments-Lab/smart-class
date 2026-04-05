import streamlit as st
import requests
import pandas as pd

if st.session_state.get("role") != "user":
    st.warning("Please log in as a user to access this page.")
    st.stop()

# Custom CSS from your notebook
st.markdown('''
<style>
.main-header {
    background-color: #ffc1cc;
    padding: 20px;
    text-align: center;
    color: #333;
    font-size: 28px;
    font-weight: 700;
    border-radius: 10px;
    margin-bottom: 2rem;
}
.to-label { text-align:center; font-size:1.4rem; font-weight:bold; color:#ff6b81; }
.stButton>button { background-color:#ff6b81; color:white; border:none; border-radius:8px; }
[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-header">Book a Classroom or Lab</div>', unsafe_allow_html=True)

# Search Form
resource_type = st.radio("Looking for a:", ["Classroom", "Lab"], horizontal=True)
selected_date = st.date_input("Select Date", value=None)

st.markdown("### Time Slot")
col_h1, col_m1, col_to, col_h2, col_m2 = st.columns([1,1,0.5,1,1])
with col_h1:
    start_hour = st.selectbox("start_h", [f"{i:02d}" for i in range(24)], label_visibility="collapsed")
with col_m1:
    start_min = st.selectbox("start_m", [f"{i:02d}" for i in range(60)], label_visibility="collapsed")
with col_to:
    st.markdown("<div class='to-label'>to</div>", unsafe_allow_html=True)
with col_h2:
    end_hour = st.selectbox("end_h", [f"{i:02d}" for i in range(24)], label_visibility="collapsed")
with col_m2:
    end_min = st.selectbox("end_m", [f"{i:02d}" for i in range(60)], label_visibility="collapsed")

purpose = st.text_input("Purpose of Booking (Class / Section)").strip()

if st.button("Search Available Slots", use_container_width=True):
    if not selected_date or not purpose:
        st.error("Please select a date and enter a purpose.")
    else:
        start_time = f"{start_hour}:{start_min}"
        end_time = f"{end_hour}:{end_min}"
        
        payload = {
            "date": str(selected_date),
            "start_time": start_time,
            "end_time": end_time,
            "resource_type": resource_type
        }
        
        with st.spinner("Searching..."):
            res = requests.post("http://127.0.0.1:8000/search", json=payload)
            
            if res.status_code == 200:
                st.session_state.search_results = res.json().get("slots", [])
                st.session_state.search_params = payload
                st.session_state.purpose = purpose
            else:
                st.error(res.json().get("detail", "Error searching slots."))
                st.session_state.search_results = None

# Display Results & Booking Logic
if st.session_state.get("search_results") is not None:
    slots = st.session_state.search_results
    
    if not slots:
        st.warning("No free slots found for the selected time window.")
    else:
        st.success(f"Found **{len(slots)}** free slot(s).")
        
        # Build display table
        data = [{"No.": i+1, "Resource": s["resource"], "Time": s["time_slot"]} for i, s in enumerate(slots)]
        st.dataframe(pd.DataFrame(data), use_container_width=True, hide_index=True)
        
        st.markdown("### Book Slots")
        options = [f"{d['No.']}. {d['Resource']} — {d['Time']}" for d in data]
        selected = st.multiselect("Choose slots to book", options=options, label_visibility="collapsed")
        
        if selected:
            if st.button(f"Book {len(selected)} Selected Slot(s)", use_container_width=True):
                # Group selected slots by resource to handle backend expectations
                # The backend expects one resource_name per request
                selected_indices = [int(opt.split(".")[0]) - 1 for opt in selected]
                
                # Group by resource
                bookings_by_resource = {}
                for idx in selected_indices:
                    slot_data = slots[idx]
                    res_name = slot_data["resource"]
                    if res_name not in bookings_by_resource:
                        bookings_by_resource[res_name] = {
                            "month_name": slot_data["month"],
                            "target_column": slot_data["target_column"],
                            "rows": [],
                            "labels": []
                        }
                    bookings_by_resource[res_name]["rows"].append(slot_data["row"])
                    bookings_by_resource[res_name]["labels"].append(slot_data["time_slot"])
                
                # Send a booking request for each resource grouped
                all_success = True
                for res_name, b_data in bookings_by_resource.items():
                    book_payload = {
                        "email": st.session_state.email,
                        "resource_type": st.session_state.search_params["resource_type"],
                        "resource_name": res_name,
                        "month_name": b_data["month_name"],
                        "target_column": b_data["target_column"],
                        "slots_to_book": b_data["rows"],
                        "time_slot_labels": b_data["labels"],
                        "date": st.session_state.search_params["date"],
                        "purpose": st.session_state.purpose
                    }
                    
                    b_res = requests.post("http://127.0.0.1:8000/book", json=book_payload)
                    if b_res.status_code != 200:
                        all_success = False
                        st.error(f"Failed to book {res_name}: {b_res.json().get('detail')}")
                
                if all_success:
                    st.success("✅ Slots booked successfully!")
                    st.session_state.search_results = None # Clear after booking
