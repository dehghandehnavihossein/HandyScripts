import os
import comtypes.client

def batch_convert_ppt_to_pdf():
    """
    Scans the current directory for PowerPoint files and converts them to PDF.
    
    This function identifies all .pptx and .ppt files in the script's directory,
    initiates a background PowerPoint instance, and saves the converted files 
    into a dedicated 'converted_pdfs' subfolder.
    
    Raises:
        COMError: If the PowerPoint application fails to initialize or access files.
    """
    
    # Identify the directory where the script resides
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(base_dir, "converted_pdfs")

    # Create the output directory if it does not exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Initialize the PowerPoint application instance
    # Note: Visible must be set to 1 to comply with PowerPoint COM automation requirements
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1 

    try:
        # Iterate through files to identify PowerPoint presentations
        for filename in os.listdir(base_dir):
            if filename.lower().endswith((".pptx", ".ppt")):
                input_path = os.path.join(base_dir, filename)
                output_filename = f"{os.path.splitext(filename)[0]}.pdf"
                output_path = os.path.join(output_dir, output_filename)
                
                print(f"Processing: {filename}...")
                
                try:
                    # Open file without UI window for better stability
                    presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
                    # 32 is the PowerPoint constant for PDF format
                    presentation.SaveAs(output_path, 32)
                    presentation.Close()
                except Exception as e:
                    print(f"Failed to convert {filename}: {e}")
                    
    finally:
        # Gracefully shut down the application
        powerpoint.Quit()
        print("Batch conversion task finished successfully.")

if __name__ == "__main__":
    batch_convert_ppt_to_pdf()
