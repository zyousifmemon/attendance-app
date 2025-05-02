import streamlit as st
import pandas as pd
from datetime import datetime
import os

# File paths for batch and section
EXCEL_FILES = {
    "2K21 CS": "data/2K21_CS_user_data.xlsx",
    "2K21 IT": "data/2K21_IT_user_data.xlsx",
    "2K22 CS": "data/2K22_CS_user_data.xlsx",
    "2K22 IT": "data/2K22_IT_user_data.xlsx",
    "2K23 CS": "data/2K23_CS_user_data.xlsx",
    "2K23 IT": "data/2K23_IT_user_data.xlsx",
    "2K24 CS": "data/2K24_CS_user_data.xlsx",
    "2K24 IT": "data/2K24_IT_user_data.xlsx"
}

# Subjects for each batch and section
SUBJECTS = {
    "2K21 CS": ["Information_Security", "Internet_of_Things", "Computer_Vision", "Entrepreneurship"],
    "2K21 IT": ["Information_Security", "Internet_of_Things", "Computer_Vision", "Digital_Image_Processing"],
    "2K22 CS": ["Graph_Theory", "Compiler_Construction", "Parallel_Distri_Comput", "MAD", "Web_Technology"],
    "2K22 IT": ["Web_Technology", "MAD", "Modeling_and_Simulation", "Enterprise_System", "SNA"],
    "2K23 CS": ["Linear_Algebra", "Operating_System", "Software_Engineering", "AI","COAL", "Theory_of_Automata" ],
    "2K23 IT": ["Linear_Algebra", "Operating_System", "Software_Engineering", "AI","COAL", "Theory_of_Automata"],
    "2K24 CS": ["DLD", "OOP","Paksitan_Studies","Data_Structures", "Mathmatical_Foundation-II","Islamic_Studies","Relevance_of_Sufism"],
    "2K24 IT": ["DLD", "OOP","Paksitan_Studies","Data_Structures", "Mathmatical_Foundation-II","Islamic_Studies","Relevance_of_Sufism"]
}

hod_password = st.secrets["passwords"]["HOD_PASSWORD"]

teacher_passwords = st.secrets["teacher_passwords"]
#st.write(teacher_passwords)

# Save attendance to Excel
def save_attendance(file_name, subject, selected_rolls, selected_date):
    selected_date_str = selected_date.strftime("%Y-%m-%d")

    if os.path.exists(file_name):
        try:
            existing_data = pd.read_excel(file_name, sheet_name=subject)
        except:
            existing_data = pd.DataFrame(columns=["Roll Number", "Subject"])
    else:
        existing_data = pd.DataFrame(columns=["Roll Number", "Subject"])

    if selected_date_str not in existing_data.columns:
        existing_data[selected_date_str] = ''

    for roll in range(1, 101):
        status = 'P' if roll in selected_rolls else 'A'

        if roll in existing_data["Roll Number"].values:
            existing_data.loc[existing_data["Roll Number"] == roll, selected_date_str] = status
        else:
            new_row = {"Roll Number": roll, "Subject": subject, selected_date_str: status}
            existing_data = pd.concat([existing_data, pd.DataFrame([new_row])], ignore_index=True)

    date_columns = sorted([col for col in existing_data.columns if col not in ["Roll Number", "Subject"]])
    existing_data = existing_data[["Roll Number", "Subject"] + date_columns]

    with pd.ExcelWriter(file_name, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
        existing_data.to_excel(writer, sheet_name=subject, index=False)

# Ensure file exists with all subject sheets
def ensure_file_exists(file_name, subjects):
    if not os.path.exists(file_name):
        with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
            for subject in subjects:
                df = pd.DataFrame({
                    "Roll Number": range(1, 101),
                    "Subject": [subject]*100
                })
                df.to_excel(writer, sheet_name=subject, index=False)
def add_logo():
   st.sidebar.image("assests/unilogo.png", width=1000)  # Adjust the width of the logo


# App main function
def app():
    
    st.header("Department of Information and Computing")
    st.subheader("Student Attendance Portal")
    st.sidebar.title("Welcome to USMS Attendance Management System")
    

    current_date = datetime.now().date()
    st.markdown(f"### 📅 Attendance Date: {current_date.strftime('%d-%m-%Y')}")
    selected_date = current_date  # Forced current date, no selection allowed
    
   # current_date = datetime.now().date()
    batch_section = st.selectbox("Select your Batch and Section:", list(SUBJECTS.keys()))
    file_name = EXCEL_FILES[batch_section]
    subject_options = SUBJECTS[batch_section]
    ensure_file_exists(file_name, subject_options)

    if not subject_options:
        st.warning("No subjects assigned to this batch/section yet.")
        return

    subject = st.selectbox("Select Subject:", subject_options)
    password = st.sidebar.text_input("Enter Subject Password", type="password")
    correct_password = teacher_passwords.get(subject)

    if password == correct_password:
        st.success("✅ Access granted to take attendance.")
        selected_date = st.date_input("Select Attendance Date", current_date)

        st.markdown("### ✅ Mark Present Students (1-100)")
        selected_rolls = []

        for i in range(1, 101):
            if st.checkbox(f"Roll No. {i}", key=f"roll_{i}"):
                selected_rolls.append(i)

        if st.button("Submit Attendance"):
            if selected_rolls:
                save_attendance(file_name, subject, selected_rolls, selected_date)
                st.success("✅ Attendance submitted successfully!")
            else:
                st.warning("⚠️ Please select at least one student!")

        with st.expander("📄 View Attendance for This Subject"):
            if os.path.exists(file_name):
                try:
                    df = pd.read_excel(file_name, sheet_name=subject)
                    st.dataframe(df)
                except:
                    st.info("No attendance records found.")
            else:
                st.info("Attendance file not found.")

    elif password:
        st.error("❌ Incorrect subject password.")

# HOD Access - Modified for date range download + View Specific Subject

add_logo()
with st.sidebar.expander("🔐 HOD Attendance Management"):
   
    hod_password_input = st.text_input("Enter HOD Password", type="password")
    if hod_password_input == hod_password:
        batch_section_hod = st.selectbox("Select your Batch and Section:", list(SUBJECTS.keys()), key="batch_section_hod")
        file_name = EXCEL_FILES.get(batch_section_hod)
        subject_options_hod = SUBJECTS[batch_section_hod]
        ensure_file_exists(file_name, subject_options_hod)

        # Select Subject to View Attendance
        selected_subject = st.selectbox("Select Subject to View Attendance:", subject_options_hod)

        # Load and display subject data
        if os.path.exists(file_name):
            try:
                df_subject = pd.read_excel(file_name, sheet_name=selected_subject)
                st.markdown(f"### 📄 Attendance for {selected_subject} ({batch_section_hod})")
                st.dataframe(df_subject)
            except Exception as e:
                st.error(f"Error reading file: {e}")
        else:
            st.info("Attendance file not found.")

        st.markdown("---")

        # Date Range Selection for Download
        start_date = st.date_input("Select Start Date", datetime.now().date(), key="hod_start_date")
        end_date = st.date_input("Select End Date", datetime.now().date(), key="hod_end_date")

        if start_date > end_date:
            st.error("🚫 Start date cannot be after end date.")
        else:
            if os.path.exists(file_name):
                try:
                    df = pd.read_excel(file_name, sheet_name=None)
                    date_filtered = {}
                    date_range = pd.date_range(start=start_date, end=end_date).strftime("%Y-%m-%d").tolist()

                    for sheet, data in df.items():
                        available_dates = [col for col in data.columns if col in date_range]
                        if available_dates:
                            filtered = data[["Roll Number", "Subject"] + available_dates]
                            date_filtered[sheet] = filtered

                    if date_filtered:
                        output_path = f"filtered_{batch_section_hod.replace(' ', '_')}_{start_date}_to_{end_date}.xlsx"
                        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                            for sheet, filtered_df in date_filtered.items():
                                filtered_df.to_excel(writer, sheet_name=sheet, index=False)

                        with open(output_path, "rb") as f:
                            st.download_button("📥 Download Filtered Attendance", f, file_name=output_path)
                    else:
                        st.info("No attendance records found between selected dates.")
                except Exception as e:
                    st.error(f"Error reading file: {e}")
            else:
                st.info("Attendance file not available.")
    elif hod_password_input:
        st.error("🚫 Incorrect HOD password.")

# Main
if __name__ == "__main__":
    app()
