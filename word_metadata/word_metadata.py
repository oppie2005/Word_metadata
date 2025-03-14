import sys
import os
from win32com import client

def update_word_metadata(file_path):
    try:
        # Create Word application object
        word = client.Dispatch("Word.Application")
        # Make Word invisible
        word.Visible = False
        
        # Open the document
        doc = word.Documents.Open(file_path)
        
        # Get filename without extension and path
        filename = os.path.splitext(os.path.basename(file_path))[0]
        
        # Access document properties
        properties = doc.BuiltInDocumentProperties
        
        # Update metadata
        properties("Title").Value = filename
        # You can add more properties here if desired
        # properties("Subject").Value = "Your Subject"
        # properties("Keywords").Value = "Your Tags"
        # properties("Author").Value = "Your Name"
        
        # Save and close
        doc.Save()
        doc.Close()
        
        print(f"Successfully updated metadata for: {filename}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        
    finally:
        # Quit Word application
        word.Quit()

if __name__ == "__main__":
    # Check if a file was dragged and dropped
    if len(sys.argv) < 2:
        print("Please drag and drop a Word document onto this script.")
    else:
        file_path = sys.argv[1]  # Get the dropped file path
        if file_path.lower().endswith(('.doc', '.docx')):
            update_word_metadata(file_path)
        else:
            print("Please drop a valid Word document (.doc or .docx)")
    
    # Keep console open to see the result
    input("Press Enter to exit...")