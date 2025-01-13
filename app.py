import streamlit as st
import os
from running_data_analysis4 import main_process
import tempfile
import uuid
import shutil


# Streamlit App Title
st.title("Running Data Analysis App")

# Sidebar for Input Directory Path or File Upload
st.sidebar.header("Directories")

# Option to upload files or specify an input directory
input_method = st.sidebar.radio("Choose input method:", ["Use Local Directory", "Upload Files"])

if input_method == "Use Local Directory":
    # Text input for local directory
    input_dir = st.sidebar.text_input(
        "Input Directory",
        value=r"C:\Users\wesjessi\Documents\running data program\MT14 raw running data"
    )
else:
    # File uploader widget
    uploaded_files = st.sidebar.file_uploader("Upload your input files", accept_multiple_files=True)

    if uploaded_files:
    # 1. Create a unique subfolder name (UUID)
    session_id = str(uuid.uuid4())
    input_dir = os.path.join("uploaded_files", session_id)
    os.makedirs(input_dir, exist_ok=True)

    # 2. Save uploaded files into this unique folder
    for uploaded_file in uploaded_files:
        with open(os.path.join(input_dir, uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())

    st.sidebar.success(f"Uploaded files saved to {input_dir}")
    
st.sidebar.write("Click **Start Processing** to generate results.")

# Start Processing Button
if st.sidebar.button("Start Processing"):
    if not os.path.exists(input_dir):
        st.error("Input directory does not exist! Please check the path.")
    else:
        try:
            # Use a persistent temporary directory
            output_dir = tempfile.mkdtemp()
            st.session_state["output_dir"] = output_dir

            # Run the processing function
            main_process(input_dir, output_dir)
            
            # CLEAN UP: remove the entire unique folder with the just-processed files
            shutil.rmtree(input_dir, ignore_errors=True)
            
            st.session_state["processed"] = True
            st.success("Processing completed! Scroll down to download the results.")
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

# Check if processing is done
if st.session_state.get("processed", False):
    st.write("Download your output files below:")

    # Get the output directory
    output_dir = st.session_state.get("output_dir")

    # Provide download buttons for each file in the temp directory
    output_files = [
        f for f in os.listdir(output_dir) if os.path.isfile(os.path.join(output_dir, f))
    ]
    for file in output_files:
        file_path = os.path.join(output_dir, file)
        with open(file_path, "rb") as f:
            file_data = f.read()
        st.download_button(
            label=f"Download {file}",
            data=file_data,
            file_name=file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
