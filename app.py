import streamlit as st
import os
import tempfile
from utils import save_uploaded_file, validate_inputs
from manual_generator import generate_manual
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="Product Manual Generator",
    page_icon="📄",
    layout="wide"
)

# Create uploads directory and subdirectories if they don't exist
upload_dirs = [
    "uploads/DAP", "uploads/HMI", "uploads/Machine_Photos", "uploads/Layout_Photos",
    "uploads/SOP", "uploads/SCADA", "uploads/Alarms", "uploads/MBOM", "uploads/EBOM",
    "uploads/Pneumatic", "uploads/Laser_Doc", "uploads/E-PLAN_Drawing", 
    "uploads/output"
]

for directory in upload_dirs:
    if not os.path.exists(directory):
        os.makedirs(directory)

# Application title and description
st.title("Product Manual Generator")
st.markdown("Generate comprehensive product manuals by filling in the form below and uploading required documents.")

# Create a form for all inputs
with st.form(key="manual_form"):
    # Basic Information Section
    st.subheader("Basic Information")
    col1, col2 = st.columns(2)
    
    with col1:
        project_name = st.text_input("Project Name *", help="Enter the name of the project")
        customer = st.text_input("Customer *", help="Enter the customer name")
    
    with col2:
        project_no = st.text_input("Project Number *", help="Enter the project number")
        manual_date = st.date_input("Manual Creation Date *", help="Select the date the manual is being created")
    
    st.markdown("### Machine Specifications")
    machine_specs = st.text_area("", help="Enter machine specifications and details")
    
    # Electrical Specifications Section
    st.subheader("Electrical Specifications")
    col1, col2 = st.columns(2)
    
    with col1:
        voltage = st.text_input("Voltage", help="Enter the voltage specifications")
        current = st.text_input("Current", help="Enter the current specifications")
    
    with col2:
        power = st.text_input("Power", help="Enter the power consumption")
        frequency = st.text_input("Frequency", help="Enter the frequency specifications")
    
    # File Uploads Section
    st.subheader("Documents")
    
    # Machine photos
    machine_photos = st.file_uploader(
        "Machine Photo(s)", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True,
        help="Upload photos of the machine (optional)"
    )
    
    # Layout photos
    layout_photos = st.file_uploader(
        "Layout Photos", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True,
        help="Upload layout photos (optional)"
    )
    
    # Document uploads
    col1, col2 = st.columns(2)
    
    with col1:
        dap_file = st.file_uploader("DAP", type=["pdf", "pptx"], help="Upload the Design Approval Package (optional)")
        sop_file = st.file_uploader("SOP", type=["pdf", "docx", "pptx"], help="Upload the Standard Operating Procedure (optional)")
        hmi_file = st.file_uploader("HMI", type=["pdf", "pptx"], help="Upload the Human-Machine Interface documentation (optional)")
        scada_file = st.file_uploader("SCADA", type=["pdf", "docx", "pptx"], help="Upload the SCADA documentation (optional)")
        alarms_file = st.file_uploader("Alarms", type=["pdf", "docx", "xlsx"], help="Upload the Alarms documentation (optional)")
    
    with col2:
        laser_doc = st.file_uploader("Laser Doc", type=["pdf", "docx"], help="Upload the Laser Documentation (optional)")
        eplan_drawing = st.file_uploader("E-PLAN Drawing", type=["pdf"], help="Upload the E-PLAN Drawing (PDF only)")
        mbom_file = st.file_uploader("MBOM", type=["xlsx"], help="Upload the Manufacturing Bill of Materials (Excel only)")
        ebom_file = st.file_uploader("EBOM", type=["xlsx"], help="Upload the Engineering Bill of Materials (Excel only)")
        pneumatic_file = st.file_uploader("Pneumatic Circuit Diagram", type=["pdf", "docx"], help="Upload the Pneumatic Circuit Diagram (optional)")
    
    # Optional Inputs Section
    st.subheader("Optional Components")
    
    col1, col2 = st.columns(2)
    
    with col1:
        robot_part_no = st.text_input("Robot Part Number", help="Enter the robot part number if applicable")
    
    with col2:
        leak_test_part_no = st.text_input("Leak Testing Part Number", help="Enter the leak testing part number if applicable")
    
    # Submit button
    submit_button = st.form_submit_button(label="Generate Manual")

# Process form submission
if submit_button:
    # Validate required inputs - only text fields are required
    required_text_inputs = {
        "Project Name": project_name,
        "Customer": customer,
        "Project Number": project_no
    }
    
    # File inputs are not required
    required_file_inputs = {}
    
    validation_result, missing_fields = validate_inputs(required_text_inputs, required_file_inputs)
    
    if not validation_result:
        st.error(f"Please fill in all required fields: {', '.join(missing_fields)}")
    else:
        # Show progress
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("Creating temporary files...")
        progress_bar.progress(10)
        
        # Prepare output directory
        output_dir = os.path.join("uploads", "output", project_name.replace(" ", "_"))
        os.makedirs(output_dir, exist_ok=True)
        
        # Save uploaded files to their respective category folders
        status_text.text("Processing uploaded files...")
        progress_bar.progress(20)
        
        # Save machine photos
        if machine_photos:
            for i, photo in enumerate(machine_photos):
                file_path = os.path.join("uploads", "Machine_Photos", f"{i+1}_{photo.name}")
                with open(file_path, "wb") as f:
                    f.write(photo.getbuffer())
        
        # Save layout photos
        if layout_photos:
            for i, photo in enumerate(layout_photos):
                file_path = os.path.join("uploads", "Layout_Photos", f"{i+1}_{photo.name}")
                with open(file_path, "wb") as f:
                    f.write(photo.getbuffer())
        
        # Save individual document files
        doc_files = {
            "DAP": dap_file,
            "SOP": sop_file,
            "HMI": hmi_file,
            "SCADA": scada_file,
            "Alarms": alarms_file,
            "MBOM": mbom_file,
            "EBOM": ebom_file,
            "Pneumatic": pneumatic_file,
            "Laser_Doc": laser_doc,
            "Electrical_Circuit_Diagram": eplan_drawing
        }
        
        for doc_type, file in doc_files.items():
            if file:
                file_path = os.path.join("uploads", doc_type, file.name)
                with open(file_path, "wb") as f:
                    f.write(file.getbuffer())
        
        # Prepare project info and electrical specifications
        project_info = {
            "project_name": project_name,
            "customer": customer,
            "project_no": project_no,
            "manual_date": manual_date.strftime("%Y-%m-%d"),
            "additional_info": machine_specs,
            "robot_part_no": robot_part_no,
            "leak_test_part_no": leak_test_part_no
        }
        
        electrical_specs = {
            "voltage": voltage,
            "current": current,
            "power": power,
            "frequency": frequency
        }
        
        status_text.text("Processing and analyzing documents...")
        progress_bar.progress(40)
        
        # Generate the manual using our new module
        try:
            status_text.text("Generating manual - This may take a few moments...")
            progress_bar.progress(60)
            
            # Generate the manual
            output_file = generate_manual(project_info, electrical_specs, output_dir)
            
            progress_bar.progress(90)
            
            # Read the generated file
            with open(output_file, "rb") as file:
                status_text.text("Manual generated successfully!")
                progress_bar.progress(100)
                
                # Provide download button
                st.success("Your product manual has been generated successfully!")
                file_data = file.read()
                
                st.download_button(
                    label="Download Product Manual",
                    data=file_data,
                    file_name=f"{project_name}_Manual.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error generating manual: {str(e)}")
            st.error("Please check that you've provided valid files and try again.")
            progress_bar.empty()
            status_text.empty()