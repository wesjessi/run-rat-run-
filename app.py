import streamlit as st
import os
from running_data_analysis4 import main_process
import tempfile

# Streamlit App Title
st.title("Running Data Analysis App")

# Sidebar for Input Directory Path
st.sidebar.header("Directories")
input_dir = st.sidebar.text_input(
    "Input Directory",
    value=r"C:\Users\wesjessi\Documents\running data program\MT14 raw running data"
)

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
