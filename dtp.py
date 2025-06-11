import os
import shutil
import glob
import win32com.client as win32
from pathlib import Path
import sys
import logging
from datetime import datetime
from tqdm import tqdm # Import tqdm for progress bar

class DocumentToPdfConverter:
    """
    Document to PDF Converter

    Converts document files in the source folder to PDF,
    or copies files with specified extensions to the destination folder.
    """
    def __init__(self, source_folder, destination_folder, target_extensions=None):
        """
        Initializes the Document to PDF Converter.
        
        Args:
            source_folder: Path to the folder containing original files.
            destination_folder: Path to the folder where converted PDF files will be saved.
            target_extensions: List of file extensions to convert (default: hwp, hwpx, pdf, ppt, pptx, doc, docx).
        """
        self.source_folder = Path(source_folder)
        self.destination_folder = Path(destination_folder)
        
        if target_extensions is None:
            self.target_extensions = ['.hwp', '.hwpx', '.pdf', '.ppt', '.pptx', '.doc', '.docx']
        else:
            self.target_extensions = [ext.lower() if ext.startswith('.') else f'.{ext.lower()}' 
                                    for ext in target_extensions]
        
        # List of file extensions to copy directly
        self.copy_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.txt', '.pdf']
        
        # Logging configuration
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(f'conversion_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log', encoding='utf-8')
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # Office application objects
        self.hwp = None
        self.excel = None
        self.powerpoint = None
        self.word = None
        
    def initialize_applications(self):
        """Initializes Office applications."""
        # Initializes Office applications.
        try:
            # Initialize Hangul Word Processor (HWP)
            self.hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
            self.hwp.XHwpWindows.Item(0).Visible = False  # Run in background
            self.hwp.RegisterModule('FilePathCheckDLL', 'FileAuto')  # Auto-approve security module
            self.logger.info("Hangul application initialization complete") # Hangul application initialization complete
        except Exception as e:
            self.logger.warning(f"Hangul application initialization failed: {e}") # Hangul application initialization failed
            
        try:
            # Initialize Excel
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.logger.info("Excel application initialization complete") # Excel application initialization complete
        except Exception as e:
            self.logger.warning(f"Excel application initialization failed: {e}") # Excel application initialization failed
            
        try:
            # Initialize PowerPoint
            self.powerpoint = win32.gencache.EnsureDispatch('PowerPoint.Application')
            self.powerpoint.Visible = False  # Change to False for background execution
            self.logger.info("PowerPoint application initialization complete") # PowerPoint application initialization complete
        except Exception as e:
            self.logger.warning(f"PowerPoint application initialization failed: {e}") # PowerPoint application initialization failed
            
        try:
            # Initialize Word
            self.word = win32.gencache.EnsureDispatch('Word.Application')
            self.word.Visible = False
            self.word.DisplayAlerts = 0
            self.logger.info("Word application initialization complete") # Word application initialization complete
        except Exception as e:
            self.logger.warning(f"Word application initialization failed: {e}") # Word application initialization failed
    
    def close_applications(self):
        """Closes all Office applications."""
        # Closes all Office applications.
        if self.hwp:
            try:
                self.hwp.Quit()
                self.logger.info("Hangul application terminated.") # Hangul application terminated.
            except:
                pass
                
        if self.excel:
            try:
                self.excel.Quit()
                self.logger.info("Excel application terminated.") # Excel application terminated.
            except:
                pass
                
        if self.powerpoint:
            try:
                self.powerpoint.Quit()
                self.logger.info("PowerPoint application terminated.") # PowerPoint application terminated.
            except:
                pass
                
        if self.word:
            try:
                self.word.Quit()
                self.logger.info("Word application terminated.") # Word application terminated.
            except:
                pass
    
    def find_target_files(self):
        """Finds files with specified extensions."""
        # Finds files with specified extensions.
        target_files = []
        
        all_extensions = self.target_extensions + self.copy_extensions
        unique_extensions = sorted(list(set(all_extensions))) # Remove duplicates and sort

        for ext in unique_extensions:
            # Recursively find files in all subfolders
            pattern = f"**/*{ext}"
            files = list(self.source_folder.glob(pattern))
            target_files.extend(files)
        
        self.logger.info(f"Found {len(target_files)} target files.") # Found {len(target_files)} target files.
        return target_files
    
    def create_directory_structure(self, source_file):
        """Creates a directory structure in the destination folder identical to the source folder."""
        # Creates a directory structure in the destination folder identical to the source folder.
        # Calculate relative path of the source file
        relative_path = source_file.relative_to(self.source_folder)
        
        # Create target path
        target_path = self.destination_folder / relative_path.parent
        
        # Create directory if it does not exist
        target_path.mkdir(parents=True, exist_ok=True)
        
        return target_path
    
    def copy_non_pdf_file(self, source_file):
        """Copies the file to the destination folder without PDF conversion."""
        # Copies the file to the destination folder without PDF conversion.
        try:
            target_dir = self.create_directory_structure(source_file)
            target_file = target_dir / source_file.name
            shutil.copy2(source_file, target_file)
            self.logger.info(f"File copy complete: {source_file.name} -> {target_file}") # File copy complete: {source_file.name} -> {target_file}
            return True
        except Exception as e:
            self.logger.error(f"Failed to copy file {source_file.name}: {e}") # Failed to copy file {source_file.name}: {e}
            return False
    
    def convert_hwp_to_pdf(self, hwp_file):
        """Converts an HWP file to PDF."""
        # Converts an HWP file to PDF.
        if not self.hwp:
            self.logger.error("Hangul application not initialized.") # Hangul application not initialized.
            return False
            
        try:
            # Create target PDF file path
            target_dir = self.create_directory_structure(hwp_file)
            pdf_file = target_dir / hwp_file.name.replace(hwp_file.suffix, '.pdf')
            
            # Open the HWP file
            self.hwp.Open(str(hwp_file))
            
            # Save as PDF
            self.hwp.HAction.GetDefault('FileSaveAsPdf', self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.filename = str(pdf_file)
            self.hwp.HParameterSet.HFileOpenSave.Format = 'PDF'
            self.hwp.HAction.Execute("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)

            self.hwp.HAction.Execute("FileClose") # Close HWP file
            
            self.logger.info(f"HWP -> PDF conversion complete: {hwp_file.name}") # HWP to PDF conversion complete.
            return True
            
        except Exception as e:
            # Check for specific error code and consider it as success
            if isinstance(e, win32.pywintypes.com_error) and e.args[0] == -2147352561:
                self.logger.warning(f"Specific warning during HWP -> PDF conversion (considered as completed): {hwp_file}: {e}") # Specific warning during HWP to PDF conversion (considered as completed).
                return True
            else:
                self.logger.error(f"HWP -> PDF conversion failed {hwp_file}: {e}") # HWP to PDF conversion failed.
                return False
    
    def convert_excel_to_pdf(self, excel_file):
        """Converts an Excel file to PDF."""
        # Converts an Excel file to PDF.
        if not self.excel:
            self.logger.error("Excel application not initialized.") # Excel application not initialized.
            return False
            
        try:
            # Create target PDF file path
            target_dir = self.create_directory_structure(excel_file)
            pdf_file = target_dir / excel_file.name.replace(excel_file.suffix, '.pdf')
            
            # Open the Excel file
            workbook = self.excel.Workbooks.Open(str(excel_file))
            
            # Save as PDF
            workbook.ExportAsFixedFormat(0, str(pdf_file))
            workbook.Close()
            
            self.logger.info(f"Excel -> PDF conversion complete: {excel_file.name}") # Excel to PDF conversion complete.
            return True
            
        except Exception as e:
            self.logger.error(f"Excel -> PDF conversion failed {excel_file.name}: {e}") # Excel to PDF conversion failed.
            return False
    
    def convert_powerpoint_to_pdf(self, ppt_file):
        """Converts a PowerPoint file to PDF."""
        # Converts a PowerPoint file to PDF.
        if not self.powerpoint:
            self.logger.error("PowerPoint application not initialized.") # PowerPoint application not initialized.
            return False
            
        try:
            # Create target PDF file path
            target_dir = self.create_directory_structure(ppt_file)
            pdf_file = target_dir / ppt_file.name.replace(ppt_file.suffix, '.pdf')
            
            # Open the PowerPoint file
            presentation = self.powerpoint.Presentations.Open(str(ppt_file))
            
            # Save as PDF (ppFixedFormatTypePDF = 2)
            presentation.SaveAs(str(pdf_file), 32)  # 32 = ppSaveAsPDF
            presentation.Close()
            
            self.logger.info(f"PowerPoint -> PDF conversion complete: {ppt_file.name}") # PowerPoint to PDF conversion complete.
            return True
            
        except Exception as e:
            self.logger.error(f"PowerPoint -> PDF conversion failed {ppt_file.name}: {e}") # PowerPoint to PDF conversion failed.
            return False
    
    def convert_word_to_pdf(self, word_file):
        """Converts a Word file to PDF."""
        # Converts a Word file to PDF.
        if not self.word:
            self.logger.error("Word application not initialized.") # Word application not initialized.
            return False
            
        try:
            # Create target PDF file path
            target_dir = self.create_directory_structure(word_file)
            pdf_file = target_dir / word_file.name.replace(word_file.suffix, '.pdf')
            
            # Open the Word file
            document = self.word.Documents.Open(str(word_file))
            
            # Save as PDF (wdExportFormatPDF = 17)
            document.SaveAs(str(pdf_file), 17) # 17 means wdFormatPDF
            document.Close()
            
            self.logger.info(f"Word -> PDF conversion complete: {word_file.name}") # Word to PDF conversion complete.
            return True
            
        except Exception as e:
            self.logger.error(f"Word -> PDF conversion failed {word_file.name}: {e}") # Word to PDF conversion failed.
            return False
    
    def convert_file_to_pdf(self, file_path):
        """Calls the appropriate conversion function based on the file extension."""
        # Calls the appropriate conversion function based on the file extension.
        file_extension = file_path.suffix.lower()
        
        # Check for extensions to copy
        if file_extension in self.copy_extensions:
            self.logger.info(f"Copying non-PDF file: {file_path.name}") # Copying non-PDF file: {file_path.name}
            return self.copy_non_pdf_file(file_path)
        
        # Check for extensions to convert to PDF
        elif file_extension == '.hwp' or file_extension == '.hwpx':
            self.logger.info(f"Converting HWP to PDF: {file_path.name}") # Converting HWP to PDF: {file_path.name}
            return self.convert_hwp_to_pdf(file_path)
        elif file_extension == '.xls' or file_extension == '.xlsx':
            self.logger.info(f"Converting Excel to PDF: {file_path.name}") # Converting Excel to PDF: {file_path.name}
            return self.convert_excel_to_pdf(file_path)
        elif file_extension == '.ppt' or file_extension == '.pptx':
            self.logger.info(f"Converting PowerPoint to PDF: {file_path.name}") # Converting PowerPoint to PDF: {file_path.name}
            return self.convert_powerpoint_to_pdf(file_path)
        elif file_extension == '.doc' or file_extension == '.docx':
            self.logger.info(f"Converting Word to PDF: {file_path.name}") # Converting Word to PDF: {file_path.name}
            return self.convert_word_to_pdf(file_path)
        else:
            self.logger.warning(f"Unsupported file type, skipping: {file_path.name}") # Unsupported file type, skipping: {file_path.name}
            return False
            
    def process_files(self):
        """Executes the entire file processing workflow."""
        # Executes the entire file processing workflow.
        try:
            # Initialize Office applications
            self.initialize_applications()
            
            # Find target files
            target_files = self.find_target_files()
            
            if not target_files:
                self.logger.info(f"No target files found in {self.source_folder}. Exiting.") # No target files found in {self.source_folder}. Exiting.
                return
                
            self.logger.info(f"Starting file processing from {self.source_folder} to {self.destination_folder}") # Starting file processing from {self.source_folder} to {self.destination_folder}
            
            success_count = 0
            fail_count = 0
            
            # Wrap the iteration with tqdm for a progress bar
            with tqdm(target_files, desc="Processing files", unit="file") as pbar:
                for file_path in pbar:
                    # Update tqdm description with current file name
                    pbar.set_description(f"Processing: {file_path.name}")
                    if self.convert_file_to_pdf(file_path):
                        success_count += 1
                    else:
                        fail_count += 1
            
            self.logger.info(f"\nProcessing complete. Total files processed: {len(target_files)}, Successful: {success_count}, Failed: {fail_count}") # Processing complete. Total files processed: {len(target_files)}, Successful: {success_count}, Failed: {fail_count}
            
        except Exception as e:
            self.logger.error(f"An error occurred during file processing: {e}") # An error occurred during file processing: {e}
        finally:
            # Close Office applications
            self.close_applications()
            self.logger.info("All applications closed. Program terminated.") # All applications closed. Program terminated.

def main():
    """Main function"""
    # Main function.
    print("=" * 60)
    print("Document PDF Converter") # Document PDF Converter
    print("=" * 60)
    
    # User input
    source_folder = input("Enter the source folder path: ").strip()
    destination_folder = input("Enter the destination folder path: ").strip()
    
    # Enter extensions to convert (optional)
    extensions_input = input("Enter extensions to convert (comma-separated, e.g., hwp,hwpx,xls,ppt) [Default: All supported formats]: ").strip()
    
    target_extensions = None
    if extensions_input:
        target_extensions = [ext.strip() for ext in extensions_input.split(',')]
    
    # Path validation
    if not os.path.exists(source_folder):
        print(f"Error: Source folder does not exist: {source_folder}") # Error: Source folder does not exist.
        return
    
    # Create and run converter
    converter = DocumentToPdfConverter(source_folder, destination_folder, target_extensions)
    
    print(f"\nSource folder: {source_folder}") # Source folder:
    print(f"Destination folder: {destination_folder}") # Destination folder:
    print(f"Target extensions for conversion: {converter.target_extensions}") # Target extensions for conversion:
    
    confirm = input("\nStart conversion? (y/n): ").strip().lower()
    if confirm == 'y':
        converter.process_files()
    else:
        print("Conversion cancelled.") # Conversion cancelled.


if __name__ == "__main__":
    # Entry point of the script.
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProgram interrupted by user.") # Program interrupted by user.
    except Exception as e:
        print(f"\nAn error occurred during program execution: {e}") # An error occurred during program execution:
        input("Press Enter to exit...") # Press Enter to exit...
