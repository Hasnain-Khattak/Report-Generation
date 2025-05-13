
import os
from datetime import datetime
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from dotenv import load_dotenv
from docx import Document
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
import logging

from  .llm_processor import LLMProcessor

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class CorrectiveActionsExtractor:
    """Extract corrective actions data from an audit report for the register."""
    
    @staticmethod
    def extract_from_report(report_content, report_date=None, score_data=None):
        """Extract key information for the corrective actions register."""
        if not report_date:
            report_date = datetime.now().strftime("%d/%m/%Y")
        
        # Initialize data dictionary with default values
        data = {
            "Date": report_date,
            "Source of Issue": "Internal Audit",  # Default, will be updated based on content
            "Type": "Opportunity for Improvement",  # Default, will try to detect
            "Details": "",
            "Root Cause": "Infancy of the quality management system",  # Default
            "Person": "Management team",  # Default
            "Corrective Actions Implemented": "",
            "Actual close out date": ""
        }
        
        # Extract data from report content
        lines = report_content.split('\n')
        
        # Enhanced extraction for process details and types
        process_details = ""
        found_ofi = False
        found_nc = False
        in_process_table = False
        process_name = ""
        issue_details = []
        
        # Determine source of issue based on report content
        source_of_issue = CorrectiveActionsExtractor.determine_source_of_issue(report_content)
        if source_of_issue:
            data["Source of Issue"] = source_of_issue
        
        for i, line in enumerate(lines):
            # Check for type classification
            if "NONCONFORMANCES" in line.upper() and i+1 < len(lines) and lines[i+1].strip() == "Yes":
                found_nc = True
                data["Type"] = "Nonconformance"
            
            elif "OPPORTUNITIES FOR IMPROVEMENTS" in line.upper() and i+1 < len(lines) and lines[i+1].strip() == "Yes":
                found_ofi = True
                data["Type"] = "Opportunity for Improvement"
            
            # Look for process table
            if "PROCESS" in line.upper() and "OK" in line.upper() and "OFI" in line.upper() and "NC" in line.upper():
                in_process_table = True
                continue
            
            # Extract details from process table
            if in_process_table and line.strip():
                # if "✓" in line and "OFI" not in line and "NC" not in line and "⨯" not in line:
                if "✓" in line and ("OFI" in line or "NC" in line):
                    # This is a passed process - not an issue
                    continue
                elif ("OFI" in line or "NC" in line or "⨯" in line):
                    # Found an issue in the process table
                    parts = line.split()
                    if len(parts) > 1:
                        process_name = parts[0]
                        
                        # Look for the additional comments section
                        for j in range(len(parts)):
                            if j > 1 and not parts[j].startswith("✓") and not parts[j] in ["OFI", "NC", "NA"]:
                                process_details = " ".join(parts[j:])
                                if process_details:
                                    issue_detail = f"Issue in {process_name} process: {process_details}"
                                    issue_details.append(issue_detail)
                                    break
            
            # Look for AUDIT REPORT FINAL COMMENTS section
            if "AUDIT REPORT FINAL COMMENTS" in line.upper():
                comment_buffer = []
                for j in range(i+1, min(i+15, len(lines))):
                    if lines[j].strip() and not lines[j].startswith("Hasnain") and not lines[j].startswith("Lead"):
                        comment_buffer.append(lines[j].strip())
                
                if comment_buffer:
                    final_comments = " ".join(comment_buffer[:3])  # Get up to 3 lines of comments
                    issue_details.append(f"Audit comments: {final_comments}")
        
        # Combine all issue details
        if issue_details:
            data["Details"] = " | ".join(issue_details)
        
        # If we still don't have details, look for any sections mentioning issues or recommendations
        if not data["Details"] or len(data["Details"]) < 10:
            for i, line in enumerate(lines):
                if any(keyword in line.upper() for keyword in ["ISSUE", "IMPROVEMENT", "RECOMMENDATION", "FINDING"]):
                    context_lines = []
                    for j in range(i, min(i+5, len(lines))):
                        if lines[j].strip():
                            context_lines.append(lines[j].strip())
                    
                    if context_lines:
                        data["Details"] = " ".join(context_lines[:2])
                        break
        
        # Generate corrective actions based on the identified details
        if data["Details"] and (not data["Corrective Actions Implemented"] or len(data["Corrective Actions Implemented"]) < 10):
            data["Corrective Actions Implemented"] = CorrectiveActionsExtractor.generate_corrective_action(data["Details"], data["Type"])
        
        # Try to enhance with LLM if the details are still minimal
        if "report_content" in locals() and len(report_content) > 100:
            try:
                enhanced_data = LLMProcessor.extract_corrective_actions(report_content)
                if enhanced_data:
                    # Only replace if LLM provided meaningful content
                    if enhanced_data.get("details") and len(enhanced_data["details"]) > 20:
                        data["Details"] = enhanced_data["details"]
                    if enhanced_data.get("corrective_actions") and len(enhanced_data["corrective_actions"]) > 20:
                        data["Corrective Actions Implemented"] = enhanced_data["corrective_actions"]
                    if enhanced_data.get("source_of_issue") and len(enhanced_data["source_of_issue"]) > 3:
                        data["Source of Issue"] = enhanced_data["source_of_issue"]
            except Exception as e:
                # Fallback to the previously extracted data if LLM fails
                print(f"LLM extraction failed: {str(e)}")
        
        # Set a fallback if we still don't have details
        if not data["Details"] or len(data["Details"]) < 10:
            data["Details"] = "Refer to full audit report for process improvement details."
            
        # Set a fallback if we still don't have corrective actions
        if not data["Corrective Actions Implemented"] or len(data["Corrective Actions Implemented"]) < 10:
            data["Corrective Actions Implemented"] = "Implement process improvements based on audit findings."
        
        if score_data:
            for file_name, data in score_data.items():
                if data['category'] in ['OFI', 'NC']:
                    issue_details.append(f"Issue with {os.path.basename(file_name)}: {data['comment']}")
                    
        return data
    
    @staticmethod
    def determine_source_of_issue(report_content):
        """
        Determine the source of the issue based on report content.
        
        Common sources:
        - Internal Audit
        - External Audit
        - Customer Complaint
        - Management Review
        - Process Monitoring
        - Employee Suggestion
        """
        report_content_lower = report_content.lower()
        
        # Check for specific keywords that indicate source
        if "external audit" in report_content_lower or "third party" in report_content_lower:
            return "External Audit"
        elif "customer complaint" in report_content_lower or "client feedback" in report_content_lower:
            return "Customer Complaint"
        elif "management review" in report_content_lower:
            return "Management Review"
        elif "employee suggestion" in report_content_lower or "staff feedback" in report_content_lower:
            return "Employee Suggestion"
        elif "monitoring" in report_content_lower and "process" in report_content_lower:
            return "Process Monitoring"
        elif "internal audit" in report_content_lower:
            return "Internal Audit"
            
        # If there's text in the AUDIT TYPE field, extract it
        lines = report_content.split('\n')
        for i, line in enumerate(lines):
            if "AUDIT TYPE" in line.upper() and i+1 < len(lines) and lines[i+1].strip():
                audit_type = lines[i+1].strip()
                if "internal" in audit_type.lower():
                    return "Internal Audit"
                elif "external" in audit_type.lower():
                    return "External Audit"
                else:
                    return audit_type.strip()
        
        # Default fallback
        return "Internal Audit"
    
    @staticmethod
    def generate_corrective_action(details, issue_type):
        """
        Generate appropriate corrective actions based on the details and issue type.
        """
        details_lower = details.lower()
        
        # Check for specific issues and provide targeted actions
        if "documentation" in details_lower or "record" in details_lower or "document" in details_lower:
            return "Revise documentation system to ensure proper maintenance and accessibility of quality records."
        
        elif "training" in details_lower or "competence" in details_lower or "knowledge" in details_lower:
            return "Implement targeted training program to address identified competency gaps."
        
        elif "customer" in details_lower or "client" in details_lower or "feedback" in details_lower:
            return "Enhance customer feedback mechanism and implement systematic review process for complaints and suggestions."
        
        elif "equipment" in details_lower or "maintenance" in details_lower or "calibration" in details_lower:
            return "Revise equipment maintenance schedule and establish verification procedures for critical equipment."
        
        elif "supplier" in details_lower or "vendor" in details_lower or "purchasing" in details_lower:
            return "Improve supplier evaluation process and implement regular performance reviews for critical suppliers."
        
        elif "process" in details_lower and ("control" in details_lower or "monitoring" in details_lower):
            return "Establish additional process controls and monitoring mechanisms to prevent recurrence."
        
        # Generate based on issue type if no specific match
        elif issue_type == "Nonconformance":
            return "Conduct root cause analysis and implement systemic changes to prevent recurrence of the nonconformity."
        
        else:  # Opportunity for Improvement
            return "Develop and implement process enhancements to address the identified opportunity for improvement."
