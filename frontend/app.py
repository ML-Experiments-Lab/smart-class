import streamlit as st
import requests

st.set_page_config(page_title="Smart Classroom Auth", layout="centered")

# Inject your custom CSS
st.markdown('''
<style>
.stApp { background-color: #fff0f5; }
.stButton>button { background-color:#ff6b81; color:white; border-radius:8px; border: none; font-weight: 600; padding: 0.5rem 1rem;}
.main-header { background-color: #ffc1cc; padding: 20px; text-align: center; border-radius: 10px; font-size: 24px; font-weight: bold; margin-bottom: 20px;}
[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-header">Smart Classroom & Lab System</div>', unsafe_allow_html=True)

# Initialize session state variables
if "role" not in st.session_state:
    st.session_state.role = None
    st.session_state.email = None

# Base URL for FastAPI backend
API_URL = "https://smart-class-api-xez6.onrender.com"

if st.session_state.role is None:
    tab1, tab2 = st.tabs(["Login", "Register"])
    
    with tab1:
        st.subheader("Login")
        log_email = st.text_input("Email", key="log_email")
        log_pass = st.text_input("Password", type="password", key="log_pass")
        
        if st.button("Login", use_container_width=True):
            if not log_email or not log_pass:
                st.warning("Please enter both email and password.")
            elif not log_email.endswith("@adaniuni.ac.in"):
                st.error("Access denied: Only @adaniuni.ac.in emails are allowed.")
            else:
                try:
                    with st.spinner("Logging in..."):
                        res = requests.post(f"{API_URL}/auth/login", json={"email": log_email, "password": log_pass})
                        data = res.json()
                        
                        if "role" in data:
                            st.session_state.role = data["role"]
                            st.session_state.email = data["email"]
                            st.success("Logged in successfully!")
                            
                            # AUTO-REDIRECT based on role
                            if data["role"] == "admin":
                                st.switch_page("pages/Admin_Panel.py")
                            else:
                                st.switch_page("pages/User_Panel.py")
                        else:
                            st.error(data.get("error", "Invalid credentials."))
                except requests.exceptions.ConnectionError:
                    st.error("Cannot connect to the backend server. Is FastAPI running on port 8000?")

    with tab2:
        st.subheader("User Registration")
        reg_email = st.text_input("Adani Uni Email", key="reg_email")
        reg_pass = st.text_input("Password", type="password", key="reg_pass")
        
        if st.button("Register", use_container_width=True):
            if not reg_email or not reg_pass:
                st.warning("Please fill out all fields.")
            elif not reg_email.endswith("@adaniuni.ac.in"):
                st.error("Registration denied: You must use an @adaniuni.ac.in email address.")
            else:
                try:
                    with st.spinner("Registering..."):
                        res = requests.post(f"{API_URL}/auth/register", json={"email": reg_email, "password": reg_pass})
                        if res.status_code == 200:
                            data = res.json()
                            if "success" in data:
                                st.success("Registered successfully! Taking you to the dashboard...")
                                
                                # AUTO-LOGIN AND REDIRECT
                                # Since only regular users register (admin is hardcoded), we set role to "user"
                                st.session_state.role = "user"
                                st.session_state.email = reg_email
                                st.switch_page("pages/User_Panel.py")
                                
                            else:
                                # This handles the "User already exists" error from the backend
                                st.error(data.get("error", "Registration failed."))
                        else:
                            st.error(f"Backend Server Error ({res.status_code}). Check your FastAPI terminal for details.")
                except requests.exceptions.ConnectionError:
                    st.error("Cannot connect to the backend server. Is FastAPI running on port 8000?")

else:
    # View shown if they navigate back to the main URL while already logged in
    st.success(f"Logged in as **{st.session_state.email}** ({st.session_state.role.capitalize()})")
    
    # Quick links back to their respective dashboards
    if st.session_state.role == "admin":
        if st.button("Go to Admin Dashboard", use_container_width=True):
            st.switch_page("pages/Admin_Panel.py")
    else:
        if st.button("Go to User Dashboard", use_container_width=True):
            st.switch_page("pages/User_Panel.py")
            
    st.markdown("---")
    
    if st.button("Logout", use_container_width=True):
        st.session_state.role = None
        st.session_state.email = None
        st.rerun()
