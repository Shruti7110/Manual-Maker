import os


def save_uploaded_file(uploaded_file, directory, filename):
    """
    Save an uploaded file to a specific location.
    
    Args:
        uploaded_file: The uploaded file object from Streamlit
        directory (str): The directory to save the file to
        filename (str): The filename to save as
    
    Returns:
        str: Path to the saved file or None if no file was uploaded
    """
    if uploaded_file is None:
        return None

    file_path = os.path.join(directory, filename)

    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return file_path


def validate_inputs(text_inputs, file_inputs):
    """
    Validate that all required inputs are provided.
    
    Args:
        text_inputs (dict): Dictionary of text input fields and values
        file_inputs (dict): Dictionary of file input fields and values
    
    Returns:
        tuple: (is_valid, missing_fields_list)
    """
    missing_fields = []

    # Check text inputs
    for field_name, value in text_inputs.items():
        if not value:
            missing_fields.append(field_name)

    # Check file inputs
    for field_name, value in file_inputs.items():
        if value is None or (isinstance(value, list) and len(value) == 0):
            missing_fields.append(field_name)

    return len(missing_fields) == 0, missing_fields
