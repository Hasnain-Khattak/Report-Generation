import streamlit as st
import os
import tempfile
from datetime import datetime
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
from dotenv import load_dotenv
from docx import Document
from openpyxl import  load_workbook
from reportlab.lib.pagesizes import letter             
import logging

from .document_processor import DocumentProcessor
from .screenshot_handler import EvidenceScreenshotHandler
from .corrective_extractor import CorrectiveActionsExtractor
from .excel_handler import ExcelHandler
from .llm_processor import LLMProcessor
from .response_processor import ResponsePreprocessor
from .template_analyzer import TemplateAnalyzer
from .report_generator import ReportGenerator


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)



load_dotenv(".env")

os.environ['OPENAI_API_KEY'] = os.getenv('OPENAI_API_KEY')

# Set page configuration
st.set_page_config(
    page_title="AI Audit Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)



# ------- Enhanced Streamlit User Interface -------
def main():
    st.title("AI Audit Report Generator")
    st.write("Generate professional internal audit reports from your evidence files.")
    
    # Sidebar for user info and settings
    with st.sidebar:
        st.subheader("Settings")
        
        ai_provider = st.selectbox(
            "Select AI Provider",
            options=["OpenAI"],
            index=0
        )
        
        if ai_provider == "OpenAI":
            model = st.selectbox(
                "Select OpenAI Model",
                options=["gpt-4o", "gpt-4o-mini"],
                index=0
            )
        
        
        auditor_name = st.text_input("Auditor Name")
        
        st.subheader("Instructions")
        st.info("""
        1. Upload your evidence files (PDFs, Word docs, images, etc.)
        2. Upload your audit report template (Word format)
        3. Enter your auditor information
        4. Click "Generate Report"
        5. Review and download your completed report
        6. Optionally update the Corrective Actions Register
        """)
    
    # Main area for file uploads and report generation
    tab1, tab2, tab3 = st.tabs(["Generate Report", "Corrective Actions Register", "About"])
    
    # Global state variables to share data between tabs
    if 'generated_report_content' not in st.session_state:
        st.session_state.generated_report_content = None
    if 'report_bytes' not in st.session_state:
        st.session_state.report_bytes = None
    if 'report_date' not in st.session_state:
        st.session_state.report_date = datetime.now().strftime("%d/%m/%Y")
    if 'register_updated' not in st.session_state:
        st.session_state.register_updated = False
    
    with tab1:
        # File uploads for evidence
        st.subheader("Upload Evidence Files")
        st.write("Upload files containing evidence for your audit (PDFs, Word docs, Excel, images, etc.)")
        evidence_files = st.file_uploader(
            "Drag and drop files here",
            type=["pdf", "docx", "doc", "xlsx", "xls", "jpg", "jpeg", "png", "txt", "dotx"],
            accept_multiple_files=True,
            key="evidence_files"
        )
        
        # Display number of files uploaded
        if evidence_files:
            st.success(f"{len(evidence_files)} evidence files uploaded")
        
        # logger.info(f"evidence file upload: {evidence_files}")
        
        # File upload for template
        st.subheader("Upload Report Template")
        st.write("Upload your audit report template (Word format)")
        template_file = st.file_uploader(
            "Drag and drop template here",
            type=["docx", "dotx"],
            key="template_file"
        )
        
        if template_file:
            st.success("Template file uploaded: " + template_file.name)
        
        # Generate report button
        generate_btn = st.button(
            "Generate Report", 
            disabled=not (evidence_files and template_file),
            use_container_width=True,
            key="generate_btn"
        )
        
        if generate_btn and evidence_files and template_file:
            # logger.info(f"First: {evidence_files}")
            progress_container = st.container()
            with progress_container:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Create temp directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    try:
                        # Save uploaded files
                        status_text.text("Processing uploaded files...")
                        evidence_paths = []
                        
                        for i, file in enumerate(evidence_files):
                            file_path = os.path.join(temp_dir, file.name)
                            with open(file_path, "wb") as f:
                                f.write(file.getbuffer())
                            evidence_paths.append(file_path)
                            progress_bar.progress((i + 1) / (len(evidence_files) + 5))
                        
                        template_path = os.path.join(temp_dir, template_file.name)
                        with open(template_path, "wb") as f:
                            f.write(template_file.getbuffer())
                        
                        # Extract text from evidence files
                        status_text.text("Extracting content from evidence files...")
                        all_evidence_text = ""

                        # logger.info(f"file path before DocumentProcessor:{evidence_paths}")

                        
                        for i, file_path in enumerate(evidence_paths):
                            text = DocumentProcessor.process_file(file_path)
                            all_evidence_text += f"\n\n--- EVIDENCE FROM {os.path.basename(file_path)} ---\n\n"
                            all_evidence_text += text
                            progress_bar.progress((len(evidence_files) + i + 1) / (len(evidence_files) * 2 + 5))
                        
                        # Generate screenshots of evidence files
                        status_text.text("Capturing screenshots from evidence files...")
                        # logger.info(f"file path after DocumentProcessor:{evidence_paths}")
                        screenshot_handler = EvidenceScreenshotHandler()
                        evidence_images = screenshot_handler.process_evidence_files(evidence_paths)
                        progress_bar.progress((len(evidence_files) * 2 + 1) / (len(evidence_files) * 2 + 5))
                        
                        # Analyze template structure
                        status_text.text("Analyzing template structure...")
                        template_structure = TemplateAnalyzer.extract_template_structure(template_path)
                        template_prompt = TemplateAnalyzer.format_template_for_prompt(template_structure)
                        progress_bar.progress((len(evidence_files) * 2 + 2) / (len(evidence_files) * 2 + 5))
                        
                        # Create preprocessor for handling evidence references
                        response_preprocessor = ResponsePreprocessor()
                        
                        # Different handling based on AI provider
                        status_text.text(f"Generating audit report with {ai_provider} AI...")
                        
                        if ai_provider.lower() == "openai":
                            # Add evidence file references to the prompt
                            evidence_names = [os.path.basename(path) for path in evidence_paths]
                            evidence_prompt = "Evidence files provided: " + ", ".join(evidence_names)
                            logger.info(f"FILES:{evidence_prompt}")
                            evidence_prompt += "\n\nPlease reference these evidence files where appropriate in the report, especially in the 'SIGHTED EVIDENCE' sections."
                            
                            # For OpenAI, use the batch processing directly with evidence and template
                            llm_response = LLMProcessor.analyze_with_model(
                                prompt=evidence_prompt, 
                                provider=ai_provider.lower(), 
                                model=model,
                                evidence_text=all_evidence_text,
                                template_structure=template_prompt,
                                auditor_name=auditor_name
                            )
                        else:
                            # For Gemini, use the traditional approach with single prompt
                            evidence_names = [os.path.basename(path) for path in evidence_paths]
                            evidence_prompt = "Evidence files provided: " + ", ".join(evidence_names)
                            evidence_prompt += "\n\nPlease reference these evidence files where appropriate in the report, especially in the 'SIGHTED EVIDENCE' sections."
                            
                            # Add evidence file names to the prompt
                            prompt = LLMProcessor.create_audit_prompt(all_evidence_text, template_prompt, auditor_name)
                            prompt += "\n\n" + evidence_prompt
                            
                            llm_response = LLMProcessor.analyze_with_model(
                                prompt=prompt,
                                provider=ai_provider.lower(), 
                                model=model
                            )
                        
                        # Process AI response
                        status_text.text("Processing AI response...")
                        processed_response = response_preprocessor.preprocess(llm_response)
                        progress_bar.progress((len(evidence_files) * 2 + 3) / (len(evidence_files) * 2 + 5))
                        
                        # Create the report generator and set evidence images
                        report_generator = ReportGenerator()
                        logger.info(len(evidence_images))
                        report_generator.set_evidence_images(evidence_images)
                        report_generator.set_preprocessor(response_preprocessor)
                                                
                        # Generate report document by filling the template
                        status_text.text("Filling template with audit results and evidence images...")
                        report_doc = report_generator.fill_template_document(template_path, processed_response, auditor_name)
                        
                        if report_doc:
                            # Save the report
                            output_path = os.path.join(temp_dir, "Completed_Audit_Report.docx")
                            report_doc.save(output_path)
                            progress_bar.progress(1.0)
                            status_text.text("Report generated successfully!")
                            
                            # Store the report data in session state for use in other tabs
                            st.session_state.generated_report_content = processed_response
                            
                            # Save the report bytes for download
                            with open(output_path, "rb") as file:
                                report_bytes = file.read()
                                st.session_state.report_bytes = report_bytes
                            
                            # Reset the register updated flag
                            st.session_state.register_updated = False
                            
                            # Display the AI-generated content for review
                            st.subheader("AI-Generated Report Content")
                            st.write("Review the content before downloading:")
                            
                            # Show in expander to save space
                            with st.expander("Show AI-Generated Content", expanded=False):
                                st.markdown(processed_response)
                            
                            # Display evidence preview
                            with st.expander("Evidence Screenshots Preview", expanded=False):
                                st.write("The following evidence screenshots were captured and added to the report:")
                                
                                # Create columns to display evidence thumbnails
                                cols = st.columns(3)  # Display 3 thumbnails per row
                                col_idx = 0
                                
                                for file_name, images in evidence_images.items():
                                    for img in images:
                                        # Create an in-memory file-like object
                                        img_bytes = io.BytesIO(img['data'])
                                        
                                        # Display in the appropriate column
                                        with cols[col_idx % 3]:
                                            st.image(img_bytes, caption=img['source'], width=200)
                                            st.write(f"Added to report under 'SIGHTED EVIDENCE'")
                                        
                                        col_idx += 1
                            
                            # Provide download button for the report
                            st.download_button(
                                label="Download Completed Report",
                                data=report_bytes,
                                file_name="Internal_Audit_Report.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                            
                            # Show preview of first page
                            st.subheader("Report Preview")
                            st.info("This is a preview of how your report will look. Download the document for the complete formatted report with evidence screenshots.")
                            
                            # Display a sample of the structure (first few sections)
                            preview_sections = []
                            section_count = 0
                            current_section = ""
                            
                            for line in processed_response.split('\n'):
                                if line.startswith('#'):
                                    if current_section and section_count < 5:
                                        preview_sections.append(current_section)
                                        section_count += 1
                                    current_section = line + "\n"
                                elif section_count < 5:
                                    current_section += line + "\n"
                            
                            if current_section and section_count < 5:
                                preview_sections.append(current_section)
                            
                            st.markdown('\n\n'.join(preview_sections))
                            
                            # Prompt to update Corrective Actions Register
                            st.info("Now you can go to the 'Corrective Actions Register' tab to update your register with the findings from this audit.")
                            
                        else:
                            status_text.text("Error generating report document.")
                            
                    except Exception as e:
                        st.error(f"Error during report generation: {str(e)}")
                        status_text.text(f"Error: {str(e)}")
    
    with tab2:
        st.subheader("Corrective Actions Register")
        
        # Check if we have a generated report
        if st.session_state.generated_report_content is None:
            st.warning("Please generate an audit report first in the 'Generate Report' tab.")
        else:
            # File upload for existing register
            st.write("Upload your existing Corrective Actions Register (optional)")
            register_file = st.file_uploader(
                "Drag and drop Excel file here (or leave empty to create a new one)",
                type=["xlsx", "xls"],
                key="register_file"
            )
            
            # Extract corrective actions from the report
            if not st.session_state.register_updated:
                extracted_data = CorrectiveActionsExtractor.extract_from_report(
                    st.session_state.generated_report_content,
                    st.session_state.report_date
                )
                
                # Show preview of extracted data
                st.subheader("Extracted Corrective Action Data")
                st.write("Review and edit this data before adding to the register:")
                
                # Allow editing of extracted data
                col1, col2 = st.columns(2)
                with col1:
                    date = st.text_input("Date", value=extracted_data["Date"])
                    source = st.text_input("Source of Issue", value=extracted_data["Source of Issue"])
                    action_type = st.selectbox(
                        "Type", 
                        options=["Nonconformance", "Opportunity for Improvement"],
                        index=0 if extracted_data["Type"] == "Nonconformance" else 1
                    )
                    details = st.text_area("Details", value=extracted_data["Details"], height=100)
                
                with col2:
                    root_cause = st.text_input("Root Cause", value=extracted_data["Root Cause"])
                    person = st.text_input("Person Responsible", value=extracted_data["Person"])
                    actions = st.text_area("Corrective Actions Implemented", value=extracted_data["Corrective Actions Implemented"], height=100)
                    close_date = st.text_input("Actual Close Out Date", value=extracted_data["Actual close out date"])
                
                # Update button
                update_btn = st.button(
                    "Update Corrective Actions Register",
                    use_container_width=True,
                    key="update_register_btn"
                )
                
                if update_btn:
                    # Prepare final data
                    final_data = {
                        "Date": date,
                        "Source of Issue": source,
                        "Type": action_type,
                        "Details": details,
                        "Root Cause": root_cause,
                        "Person": person,
                        "Corrective Actions Implemented": actions,
                        "Actual close out date": close_date
                    }
                    
                    # Add to register
                    try:
                        updated_register = ExcelHandler.add_action_to_register(register_file, final_data)
                        register_bytes = ExcelHandler.save_register_to_bytes(updated_register)
                        
                        # Mark as updated
                        st.session_state.register_updated = True
                        
                        st.success("Corrective Actions Register updated successfully!")
                        
                        # Provide download button
                        st.download_button(
                            label="Download Updated Corrective Actions Register",
                            data=register_bytes,
                            file_name="Corrective_Actions_Register.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error updating register: {str(e)}")
            else:
                st.success("Corrective Actions Register has been updated.")
                st.info("Generate a new report to add more corrective actions.")
    
    with tab3:
        st.subheader("About this Application")
        st.write("""
        ### AI Audit Report Generator
        
        This application uses artificial intelligence to generate professional internal audit reports based on your evidence files and report templates.
        
        #### Key Features:
        - **Template Preservation**: Maintains your original template format and structure
        - **Multi-format Support**: Process PDFs, Word documents, Excel spreadsheets, images, and text files
        - **Advanced AI Analysis**: Leverages OpenAI (GPT-4) or Google (Gemini) models for expert-level audit analysis
        - **Professional Output**: Creates well-formatted, professional documents ready for stakeholders
        - **ISO Standards**: Follows ISO audit practices and terminology
        - **Corrective Actions Register**: Automatically updates your corrective actions tracking spreadsheet
        
        #### How it Works:
        1. We extract text from your evidence files using specialized processors for each file type
        2. We analyze your template to understand its structure
        3. The AI evaluates the evidence against ISO audit standards
        4. The original template is filled with the generated content
        5. Key findings are extracted to update your Corrective Actions Register
        6. You can review and download both the completed report and updated register
        
        #### Privacy & Security:
        All processing happens within this application. Your files are not stored permanently and are deleted after processing.
        """)
