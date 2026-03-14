import streamlit as st
import os
import subprocess
import smtplib
from email.message import EmailMessage

st.set_page_config(
    page_title="Parcon BOP Dashboard",
    page_icon="🍃",
    layout="wide")

# ---------------- PAGE DESIGN ---------------- #

st.markdown("""
<style>
.block-container{
    padding-top: 1rem;
}

.main-title{
    text-align:center;
    color:#2E86C1;
    font-size:42px;
    font-weight:700;
    margin-top:25px;
    margin-bottom:7px;
}

.sub-title{
    text-align:center;
    font-size:28px;
}

.title-line{
    border-bottom:2px solid #b0b0b0;
    width:60%;
    margin:auto;
    margin-top:10px;
}

.email-box{
    background-color:#f8fbff;
    padding:20px;
    border-radius:10px;
    border:1px solid #d6e4f0;
    width:40%;
    margin:auto;
    text-align:center;
    margin-top:20px;
    margin-bottom:25px;
}

.email-title{
    font-size:20px;
    font-weight:600;
    color:#2E86C1;
    margin-bottom:10px;
}
</style>

<div class="main-title">Tea CIP (Commodity Intelligence Platform)</div>
<div class="sub-title">EST / BLF Batting Order Position</div>
<div class="title-line"></div>
""", 
unsafe_allow_html=True)


# ---------------- SETTINGS ---------------- #

working_directory = r"D:\Oushnik Sarkar\Python\BATTING ORDER"

modules = [
    {"name":"AS_EST","script":"AS_EST.py","output":"AS_EST.xlsx"},
    {"name":"AS_BLF","script":"AS_BLF.py","output":"AS_BLF.xlsx"},
    {"name":"DO_TR_EST","script":"DO.TR_EST.py","output":"DO_TR_EST.xlsx"},
    {"name":"DO_TR_BLF","script":"DO.TR_BLF.py","output":"DO_TR_BLF.xlsx"},
    {"name":"CATP","script":"CA.TP.py","output":"CATP.xlsx"},
    {"name":"AS_ORTH","script":"AS_ORTH.py","output":"AS_ORTH.xlsx"},
    {"name":"AS_ORTH_EST","script":"AS_ORTH_EST.py","output":"AS_ORTH_EST.xlsx"},
    {"name":"AS_ORTH_BLF","script":"AS_ORTH_BLF.py","output":"AS_ORTH_BLF.xlsx"}
]

combined_output = "EST BLF BATTING ORDER UPTO SALE 10_updated.xlsx"

# ---------------- EMAIL FUNCTION ---------------- #

#receiver_email = st.text_input("Enter Receiver Email", placeholder="xxxxx@email.com")

#st.markdown('<div class="email-box">', unsafe_allow_html=True)
st.markdown('<div class="email-title">✉️ Enter Your Email Address</div>', unsafe_allow_html=True)

receiver_email = st.text_input(
    "Receiver Email",
    placeholder="xxxxx@email.com",
    label_visibility="collapsed"
)

st.markdown('</div>', unsafe_allow_html=True)

def send_email(file_path, receiver_email):

    sender_name = "Oushnik Sarkar"
    sender_email = "website@parcon.in"
    receiver_name = "User"
    #receiver_email = "oushnik@gmail.com"

    msg = EmailMessage()
    msg["Subject"] = "EST BLF BATTING ORDER"
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = receiver_email

    msg.set_content("""Dear Sir,

Please find the attached file.

Regards
Oushnik
""")

    with open(file_path,"rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(file_path)
        )

    with smtplib.SMTP("smtp.gmail.com",587) as server:
        server.starttls()
        server.login(sender_email,"xusq bocs tgrk kwig")
        server.send_message(msg)

    st.success(f"Email Sent to {receiver_email} Successfully")

# ---------------- RUN SCRIPT ---------------- #

def run_script(module):

    script_path = os.path.join(working_directory,module["script"])
    output_path = os.path.join(working_directory,module["output"])

    try:
        st.info(f"Running {module['script']}...")
        subprocess.run(["python",script_path],check=True)

        st.success(f"✅ {module['script']} completed successfully!")

        st.session_state[module["name"]] = output_path

    except:
        st.error(f"❌ Error running {module['script']}")

# ---------------- MODULE SECTION ---------------- #

st.markdown("### ⚙️ Individual Modules")

rows = [st.columns(4), st.columns(4)]

index = 0

for row in rows:
    for col in row:

        if index < len(modules):

            module = modules[index]

            with col:

                if st.button(f"🍃 {module['name']}", use_container_width=True):
                    run_script(module)

                # show download/email if already executed
                if module["name"] in st.session_state:

                    output_file = st.session_state[module["name"]]

                    if os.path.exists(output_file):

                        with open(output_file,"rb") as f:

                            st.download_button(
                                label=f"📥 Download {module['output']}",
                                data=f,
                                file_name=module["output"]
                            )

                        if st.button(f"📧 Send {module['output']} Email", key=module["name"]+"_email"):
                            #send_email(output_file)
                            if receiver_email:
                                send_email(output_file, receiver_email)
                            else:
                                st.warning("Please enter email address first")

        index += 1

# ---------------- COMBINED PROCESS ---------------- #

st.markdown("---")

if st.button("▶ Run Batting Order Process"):

    with st.spinner("Processing..."):

        for module in modules:
            script_path = os.path.join(working_directory,module["script"])
            subprocess.run(["python",script_path])

    st.success("All Scripts Completed")

# Download combined
if os.path.exists(combined_output):

    with open(combined_output,"rb") as f:

        st.download_button(
            label="📥 Download Excel File",
            data=f,
            file_name=combined_output)

# Email combined
if st.button("✉️ Send Email"):
    #send_email(combined_output)
    if receiver_email:
        send_email(combined_output, receiver_email)
    else:
        st.warning("Please enter email address first")
