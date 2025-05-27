import os
import tempfile
import shutil
from docx import Document
from datetime import datetime


from pptx_data_processing import (
    process_dap_to_docx,
    process_sop_to_docx,
    process_hmi_to_docx,
    process_scada_to_docx,
)

from Img_expraction import (
    process_machine_photos,
    process_layout_photos,
    remove_unused_placeholders,
    get_images_from_folder,
    process_pneumatic_photos
)

from pdf_doc_extractor import (
    pdf_to_images,
    insert_pdf_images_at_placeholder,
    process_eplan_pdf_to_docx,
    process_alarms_pdf_to_docx
)

from project_details import (
    extract_project_info,
    insert_project_info,
    insert_machine_specifications,
    insert_electrical_specifications
)

# Define upload directories to clean after processing
UPLOAD_DIRS = [
    "uploads/DAP", 
    "uploads/HMI", 
    "uploads/Machine_Photos", 
    "uploads/Layout_Photos",
    "uploads/SOP", 
    "uploads/SCADA", 
    "uploads/Alarms", 
    "uploads/MBOM", 
    "uploads/EBOM",
    "uploads/Pneumatic", 
    "uploads/output"
    "uploads/Laser_Doc", 
    "uploads/Electrical_Circuit_Diagram"
    "uploads/E-PLAN_Drawing",
    "uploads/Project_info",
    "uploads/Template",
    "dap_img_extracted",
    "sop_img_extracted",
    "hmi_img_extracted",
    "scada_img_extracted",
    "alarms_img_extracted",
    "eplan_img_extracted",
]

def clean_upload_directories():
    """
    Clean all uploaded files from the upload directories after document generation.
    """
    for directory in UPLOAD_DIRS:
        if os.path.exists(directory):
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f"Error cleaning up file {file_path}: {e}")


def generate_manual(output_dir):
    """
    Generate a manual by processing all files in the uploads directory.
    
    Args:
        project_info (dict): Dictionary containing project information
        electrical_specs (dict): Dictionary containing electrical specifications
        output_dir (str): Directory to save the output file
    
    Returns:
        str: Path to the generated manual file
    """
    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Load the base file
    source_file_path = "template/base_file.docx"    
    base_file_path = "uploads/Template/base_file.docx"
    if not os.path.exists(base_file_path):
        raise FileNotFoundError(f"Base file not found: {base_file_path}")
    
    shutil.copyfile(source_file_path, base_file_path)

    doc = Document(base_file_path)
    
    try:
        # Process each type of file
        print("Processing machine photos...")
        process_machine_photos(base_file_path) 
        process_machine_photos(base_file_path) 
        
        print("Adding project details...")
        insert_project_info(base_file_path)
        
        print("Processing layout photos...")
        process_layout_photos(base_file_path)

        print("Processing DAP file...")
        process_dap_to_docx(
            folder_path="uploads/DAP",
            template_path= base_file_path,
        )
        
        print("Processing SOP file...")
        process_sop_to_docx(
            folder_path="uploads/SOP",
            template_path=base_file_path,
        )

        print("Processing HMI file...")
        process_hmi_to_docx(
            folder_path="uploads/HMI",
            template_path=base_file_path,
        )

        print("Processing SCADA file...")
        process_scada_to_docx(
            folder_path="uploads/SCADA",
            template_path=base_file_path,
        )

        print("Processing alarms file...")
        process_alarms_pdf_to_docx(
            folder_path="uploads/Alarms",
            template_path=base_file_path,
        )

        print("Processing electrical circuit diagram...")
        process_eplan_pdf_to_docx(
            folder_path="uploads/E-PLAN_Drawing",
            template_path=base_file_path,
        )

        print("Processing pneumatic circuit diagram...")
        process_pneumatic_photos(doc, base_file_path)

        print("Adding Machine Specifications...")
        insert_machine_specifications(base_file_path)

        print("Adding electrical specifications...")
        insert_electrical_specifications(base_file_path)

        print("Removing unused placeholders...")
        remove_unused_placeholders(base_file_path)

        # Save the final document
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        info = {"project_name": ""}
        project_info = extract_project_info(txt_path="uploads/Project_info/Project_info.txt", info=info)
        output_file = os.path.join(output_dir, f"{project_info['project_name']}_{timestamp}.docx")
        doc.save(output_file)
        
        # Clean up uploaded files after successful document generation
        print("Cleaning up uploaded files...")
        clean_upload_directories()
        
        return output_file
        
    except Exception as e:
        print(f"Error generating manual: {str(e)}")
        raise
    
generate_manual("output/manuals")
    
    