import sys
import streamlit as st
import os
import tempfile
from datetime import datetime
import pandas as pd
import docx
from docx import Document
import comtypes.client
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import fitz  # PyMuPDF
import io
import win32gui
import win32ui
import win32con
import win32com.client
from dotenv import load_dotenv
import pytesseract
import mammoth
import base64
import win32com.client
import pythoncom
import time
from PIL import Image
import openai
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl import Workbook, load_workbook
import re
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from docxcompose.composer import Composer
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# from trial import _extract_from_excel as EXTRACT_FROM_EXCEL
                

from typing import List, Dict, Tuple, Union, Optional
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)



load_dotenv(".env")

os.environ['OPENAI_API_KEY'] = os.getenv('OPENAI_API_KEY')

# Set page configuration
# st.set_page_config(
#     page_title="AI Audit Report Generator",
#     page_icon="📊",
#     layout="wide",
#     initial_sidebar_state="expanded"
# )





# ------ Response Processor -------------
class ResponsePreprocessor:
    """Clean and format AI responses before document generation with enhanced score detection."""
    
    def __init__(self):
        """Initialize the ResponsePreprocessor."""
        self.evidence_images = {}
        self.evidence_metadata = {}  # Store score data, company names, dates
        self.image_paths = []
        self.static_text = {}
        self.evidence_replacements = {}
        self.score_data = {}  # Store scores for each evidence file
    
    def set_evidence_images(self, evidence_images, evidence_metadata=None):
        """
        Set the evidence images and metadata to be used during preprocessing.
        
        Args:
            evidence_images: Dictionary of evidence files and their images
            evidence_metadata: Dictionary with extracted scores, company names, dates
        """
        self.evidence_images = evidence_images
        self.evidence_metadata = evidence_metadata or {}
    
    def set_static_text(self, static_text):
        """
        Set static predefined text for sections like AUDIT TITLE, AUDIT SCOPE, etc.
        
        Args:
            static_text: Dictionary with section names as keys and content as values
        """
        self.static_text = static_text
    
    @staticmethod
    def clean_response(response_text):
        """Clean up the AI response to ensure proper formatting."""
        
        # Fix common formatting issues with tables
        cleaned_text = []
        lines = response_text.split('\n')
        in_table = False
        table_header_detected = False
        
        for i, line in enumerate(lines):
            # Detect start of table
            if line.strip().startswith('|') and line.strip().endswith('|'):
                if not in_table:
                    in_table = True
                    # If this is the first line of the table, ensure it's properly formatted
                    if '|' in line and not table_header_detected:
                        table_header_detected = True
                
                # Ensure the line has properly formatted pipe separators
                cells = line.split('|')
                cleaned_line = '|' + '|'.join(cell.strip() for cell in cells[1:-1]) + '|'
                cleaned_text.append(cleaned_line)
                
                # Check if the next line is the separator row
                if i + 1 < len(lines) and '-+-' in lines[i+1].replace('|', '+'):
                    # This is a header row, next line should be a separator
                    continue
                elif i + 1 < len(lines) and in_table and '---' in lines[i+1] and '|' in lines[i+1]:
                    # This is a proper separator row, keep it
                    continue
                elif in_table and table_header_detected and i + 1 < len(lines) and not '---' in lines[i+1] and '|' not in lines[i+1]:
                    # We need to add a separator row after the header
                    cells = line.split('|')
                    separator = '|' + '|'.join(['---' for _ in cells[1:-1]]) + '|'
                    cleaned_text.append(separator)
            
            # Detect end of table
            elif in_table and not line.strip().startswith('|'):
                in_table = False
                table_header_detected = False
                cleaned_text.append('')  # Add blank line after table
                if line.strip():  # Add non-empty line
                    cleaned_text.append(line)
            else:
                # Regular line
                cleaned_text.append(line)
        
        return '\n'.join(cleaned_text)

    @staticmethod
    def ensure_proper_headings(response_text):
        """Ensure headings use proper markdown format."""
        lines = response_text.split('\n')
        cleaned_lines = []
        
        for line in lines:
            # Check for potential headings that don't use markdown format
            if re.match(r'^[A-Z][A-Za-z\s]+:$', line.strip()):
                # Convert "Heading:" format to "# Heading"
                heading_text = line.strip().rstrip(':')
                cleaned_lines.append(f"## {heading_text}")
            else:
                cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    @staticmethod
    def standardize_table_format(response_text):
        """Standardize table formats across different AI providers' outputs."""
        # Find all tables in the response
        table_pattern = r'(\|.*\|[\r\n]+\|[-\s:|]+\|[\r\n]+((?:\|.*\|[\r\n]+)*))'
        tables = re.finditer(table_pattern, response_text, re.MULTILINE)
        
        result = response_text
        for table_match in tables:
            if not isinstance(table_match, type(None)):
                table = table_match.group(0)
                # Ensure separator row uses standard format
                lines = table.split('\n')
                if len(lines) >= 2:
                    header_line = lines[0]
                    separator_line = lines[1]
                    
                    # Count columns in header
                    column_count = header_line.count('|') - 1
                    
                    # Create standardized separator
                    standard_separator = '|' + '|'.join(['---' for _ in range(column_count)]) + '|'
                    
                    # Replace the separator line
                    standardized_table = table.replace(separator_line, standard_separator)
                    result = result.replace(table, standardized_table)
        
        return result
    
    @staticmethod
    def fix_checkmark_symbols(response_text):
        """Standardize checkmark symbols used in the report."""
        # Common variations of checkmarks from different LLMs
        checkmark_variations = ['✓', '✔', '✔️', 'X', 'x', '✗', '✘', '☑', '☒', '☐']
        
        # Standardize to a single checkmark symbol
        result = response_text
        for symbol in checkmark_variations:
            if symbol != '✓':  # We're standardizing to this symbol
                result = result.replace(f' {symbol} ', ' ✓ ')
                result = result.replace(f'|{symbol}|', '|✓|')
                result = result.replace(f'| {symbol} |', '| ✓ |')
                # Handle cases where it's just the symbol alone in a cell
                result = result.replace(f'| {symbol} ', '| ✓ ')
                result = result.replace(f' {symbol} |', ' ✓ |')
        
        # Special case for OK/OFI/NC/NA columns
        lines = result.split('\n')
        for i, line in enumerate(lines):
            if '|' in line:
                cells = line.split('|')
                for j, cell in enumerate(cells):
                    cell_stripped = cell.strip()
                    # If a cell just contains a symbol in the OK column
                    if cell_stripped in checkmark_variations and j >= 1:
                        cells[j] = ' ✓ '
                lines[i] = '|'.join(cells)
        
        return '\n'.join(lines)
    
    def analyze_evidence_scores(self):
        """
        Analyze scores in evidence to determine OK/OFI/NC status and generate comments.
        
        This is a key enhancement to determine which column should have the checkmark.
        """
        score_analysis = {}
        
        for file_name, metadata in self.evidence_metadata.items():
            if not metadata or 'scores' not in metadata or not metadata['scores']:
                # Default to OK for Excel files
                if file_name.lower().endswith(('.xlsx', '.xls')):
                    score_analysis[file_name] = {
                        'category': 'OK',
                        'comment': ""
                    }
                else:
                    score_analysis[file_name] = {
                        'category': 'NA',
                        'comment': ""
                    }
                continue
            
            # Process scores
            for score in metadata['scores']:
                # Check for patterns like X/Y
                if '/' in score:
                    try:
                        parts = score.split('/')
                        numerator = float(parts[0].strip())
                        denominator = float(parts[1].strip())
                        
                        # Calculate percentage
                        percentage = (numerator / denominator) * 100
                        
                        # Determine category based on percentage
                        if percentage >= 90:
                            score_analysis[file_name] = {
                                'category': 'OK',
                                'comment': ""
                            }
                        elif percentage >= 70:
                            score_analysis[file_name] = {
                                'category': 'OK',
                                'comment': ""
                            }
                        elif percentage >= 50:
                            score_analysis[file_name] = {
                                'category': 'OFI',
                                'comment': f"The score of {score} indicates room for improvement in customer satisfaction. Consider implementing additional feedback mechanisms and follow-up processes."
                            }
                        else:
                            score_analysis[file_name] = {
                                'category': 'NC',
                                'comment': f"The low score of {score} represents a significant gap in customer satisfaction that requires immediate corrective action and root cause analysis."
                            }
                        
                        # Break after finding the first valid score
                        break
                        
                    except Exception:
                        # If conversion fails, continue to next score
                        continue
                
                # Check for direct numeric scores
                elif score.replace('.', '', 1).isdigit():
                    try:
                        value = float(score)
                        
                        # Different logic based on scale
                        if value <= 10:  # Likely 0-10 scale
                            if value >= 8:
                                score_analysis[file_name] = {
                                    'category': 'OK',
                                    'comment': ""
                                }
                            elif value >= 5:
                                score_analysis[file_name] = {
                                    'category': 'OFI',
                                    'comment': f"The rating of {value}/10 suggests areas for service improvement. Review customer comments for specific enhancement opportunities."
                                }
                            else:
                                score_analysis[file_name] = {
                                    'category': 'NC',
                                    'comment': f"The rating of {value}/10 indicates significant customer dissatisfaction requiring immediate investigation and corrective action."
                                }
                        
                        elif value <= 100:  # Likely percentage
                            if value >= 80:
                                score_analysis[file_name] = {
                                    'category': 'OK',
                                    'comment': ""
                                }
                            elif value >= 60:
                                score_analysis[file_name] = {
                                    'category': 'OFI',
                                    'comment': f"The satisfaction score of {value}% shows moderate customer satisfaction. Implement targeted improvements based on feedback analysis."
                                }
                            else:
                                score_analysis[file_name] = {
                                    'category': 'NC',
                                    'comment': f"The satisfaction score of {value}% indicates poor customer experience. Immediate process review and corrective action is required."
                                }
                        
                        # Break after finding the first valid score
                        break
                        
                    except Exception:
                        continue
            
            # Default if no scores were properly analyzed
            if file_name not in score_analysis:
                # Provide default values
                score_analysis[file_name] = {
                    'category': 'OK',  # Default to OK
                    'comment': ""
                }
        
        self.score_data = score_analysis
        return score_analysis
    
    def process_evidence_references(self, response_text):
        """
        Process any evidence references in the response and prepare them for inclusion.
        
        This method looks for evidence file references in the text and marks them
        for replacement with actual images during document generation.
        """
        if not self.evidence_images:
            return response_text
        
        # First analyze scores to determine which category each evidence falls into
        self.analyze_evidence_scores()
            
        # Look for evidence references in the SIGHTED EVIDENCE column of tables
        lines = response_text.split('\n')
        in_table = False
        is_process_table = False
        header_row = []
        evidence_col_idx = -1
        process_col_idx = -1
        ok_col_idx = -1
        ofi_col_idx = -1
        nc_col_idx = -1
        na_col_idx = -1
        comments_col_idx = -1
        
        # Create a mapping for evidence replacements and score annotations
        evidence_replacements = {}
        score_annotations = {}
        
        for i, line in enumerate(lines):
            # Detect start of table
            if line.strip().startswith('|') and line.strip().endswith('|') and not in_table:
                in_table = True
                # Parse header row
                header_cells = [cell.strip().lower() for cell in line.split('|')[1:-1]]
                header_row = header_cells
                
                # Identify if this is the process-evidence table by checking for key columns
                if 'process' in header_cells and 'sighted evidence' in header_cells:
                    is_process_table = True
                    process_col_idx = header_cells.index('process')
                    evidence_col_idx = header_cells.index('sighted evidence')
                    
                    # Find OK/OFI/NC/NA columns
                    if 'ok' in header_cells:
                        ok_col_idx = header_cells.index('ok')
                    if 'ofi' in header_cells:
                        ofi_col_idx = header_cells.index('ofi')
                    if 'nc' in header_cells:
                        nc_col_idx = header_cells.index('nc')
                    if 'na' in header_cells:
                        na_col_idx = header_cells.index('na')
                    if 'additional comments' in header_cells:
                        comments_col_idx = header_cells.index('additional comments')
            
            # Process table rows if we're in a process-evidence table
            elif in_table and is_process_table and line.strip().startswith('|'):
                # Skip separator row
                if all(cell.strip() == '' or all(c in '-:' for c in cell) for cell in line.split('|')[1:-1]):
                    continue
                
                cells = [cell.strip() for cell in line.split('|')[1:-1]]
                if len(cells) > evidence_col_idx:
                    evidence_text = cells[evidence_col_idx]
                    process_text = cells[process_col_idx] if process_col_idx >= 0 and process_col_idx < len(cells) else ""
                    
                    # Match this row to an evidence file
                    matched_file = None
                    for evidence_file in self.evidence_images:
                        # Check for evidence filename in the evidence text
                        if evidence_file.lower() in evidence_text.lower() or any(word in evidence_text.lower() for word in evidence_file.lower().split('.')):
                            matched_file = evidence_file
                            break
                        
                        # Also check if process name matches the file name
                        evidence_keyword = os.path.basename(evidence_file).split('.')[0].lower()
                        if evidence_keyword in process_text.lower():
                            matched_file = evidence_file
                            break
                    
                    if matched_file:
                        # Mark this cell for image replacement
                        evidence_replacements[i] = {
                            'row_index': i,
                            'evidence_col': evidence_col_idx,
                            'file_name': matched_file
                        }
                        
                        # Determine where to place checkmark based on score analysis
                        if matched_file in self.score_data:
                            category = self.score_data[matched_file]['category']
                            comment = self.score_data[matched_file]['comment']
                            
                            # Store checkmark placement and comments
                            score_annotations[i] = {
                                'row_index': i,
                                'category': category,
                                'ok_col': ok_col_idx,
                                'ofi_col': ofi_col_idx,
                                'nc_col': nc_col_idx,
                                'na_col': na_col_idx,
                                'comments_col': comments_col_idx,
                                'comment': comment
                            }
            
            # Detect end of table
            elif in_table and not line.strip().startswith('|'):
                in_table = False
                is_process_table = False
                header_row = []
                evidence_col_idx = -1
                process_col_idx = -1
                ok_col_idx = -1
                ofi_col_idx = -1
                nc_col_idx = -1
                na_col_idx = -1
                comments_col_idx = -1
        
        # Store the references for later use during document generation
        self.evidence_replacements = evidence_replacements
        self.score_annotations = score_annotations
        
        # Now modify the response text to add checkmarks and comments
        modified_lines = lines.copy()
        
        # Add checkmarks based on score analysis
        for row_idx, annotation in score_annotations.items():
            row_cells = modified_lines[row_idx].split('|')
            
            # Clear any existing checkmarks in OK/OFI/NC/NA columns
            for col_idx in [annotation['ok_col'], annotation['ofi_col'], annotation['nc_col'], annotation['na_col']]:
                if col_idx >= 0 and col_idx + 1 < len(row_cells):
                    row_cells[col_idx + 1] = ' '
            
            # Add checkmark to appropriate column
            if annotation['category'] == 'OK' and annotation['ok_col'] >= 0:
                row_cells[annotation['ok_col'] + 1] = ' ✓ '
            elif annotation['category'] == 'OFI' and annotation['ofi_col'] >= 0:
                row_cells[annotation['ofi_col'] + 1] = ' ✓ '
            elif annotation['category'] == 'NC' and annotation['nc_col'] >= 0:
                row_cells[annotation['nc_col'] + 1] = ' ✓ '
            elif annotation['category'] == 'NA' and annotation['na_col'] >= 0:
                row_cells[annotation['na_col'] + 1] = ' ✓ '
            
            # Add comment if needed
            if annotation['category'] in ['OFI', 'NC'] and annotation['comment'] and annotation['comments_col'] >= 0:
                row_cells[annotation['comments_col'] + 1] = f" {annotation['comment']} "
            
            # Reassemble the row
            modified_lines[row_idx] = '|'.join(row_cells)
        
        # Determine if we have any OFIs or NCs
        has_ofi = any(ann['category'] == 'OFI' for ann in score_annotations.values())
        has_nc = any(ann['category'] == 'NC' for ann in score_annotations.values())
        
        # Update NONCONFORMANCES and OPPORTUNITIES FOR IMPROVEMENTS sections
        for i, line in enumerate(modified_lines):
            if "NONCONFORMANCES" in line.upper() and i + 1 < len(modified_lines):
                if has_nc:
                    # Add Yes on the next line after NONCONFORMANCES
                    modified_lines[i+1] = "Yes"
                    # Add explanation if not already present
                    if i + 2 < len(modified_lines) and not modified_lines[i+2].strip():
                        nc_comments = [ann['comment'] for ann in score_annotations.values() if ann['category'] == 'NC']
                        if nc_comments:
                            modified_lines[i+2] = nc_comments[0]
                else:
                    # Add No on the next line after NONCONFORMANCES
                    modified_lines[i+1] = "No"
            
            if "OPPORTUNITIES FOR IMPROVEMENTS" in line.upper() and i + 1 < len(modified_lines):
                if has_ofi:
                    # Add Yes on the next line after OPPORTUNITIES FOR IMPROVEMENTS
                    modified_lines[i+1] = "Yes"
                    # Add explanation if not already present
                    if i + 2 < len(modified_lines) and not modified_lines[i+2].strip():
                        ofi_comments = [ann['comment'] for ann in score_annotations.values() if ann['category'] == 'OFI']
                        if ofi_comments:
                            modified_lines[i+2] = ofi_comments[0]
                else:
                    # Add No on the next line after OPPORTUNITIES FOR IMPROVEMENTS
                    modified_lines[i+1] = "No"
        
        # Return the modified text
        return '\n'.join(modified_lines)
    
    @staticmethod
    def extract_headers_from_content(response_text):
        """Extract section headers from the content to help with template filling."""
        headers = {}
        current_header = None
        current_content = []
        
        lines = response_text.split('\n')
        for line in lines:
            # Check for markdown headers
            if line.strip().startswith('#'):
                # Save previous header content if exists
                if current_header and current_content:
                    headers[current_header] = '\n'.join(current_content)
                    current_content = []
                
                # Extract new header text
                header_text = line.strip().lstrip('#').strip()
                current_header = header_text
            # Check for plain text headers that end with colon
            elif line.strip() and not line.strip().startswith('|') and line.strip().endswith(':'):
                # Save previous header content if exists
                if current_header and current_content:
                    headers[current_header] = '\n'.join(current_content)
                    current_content = []
                
                # Extract new header text
                header_text = line.strip().rstrip(':').strip()
                current_header = header_text
            elif current_header:
                # Collect content for current header
                current_content.append(line)
        
        # Save the last section
        if current_header and current_content:
            headers[current_header] = '\n'.join(current_content)
        
        return headers
    
    def extract_customer_data(self):
        """
        Extract customer names, dates, and scores from evidence metadata.
        This helps with populating process cells with relevant information.
        """
        customer_data = {}
        
        for file_name, metadata in self.evidence_metadata.items():
            if not metadata:
                continue
                
            data = {
                'company': None,
                'date': None,
                'score': None,
                'comments': []
            }
            
            # Extract company name
            if 'companies' in metadata and metadata['companies']:
                data['company'] = metadata['companies'][0]
                
            # Extract date
            if 'dates' in metadata and metadata['dates']:
                data['date'] = metadata['dates'][0]
                
            # Extract score
            if 'scores' in metadata and metadata['scores']:
                data['score'] = metadata['scores'][0]
                
            # Extract comments
            if 'comments' in metadata and metadata['comments']:
                data['comments'] = metadata['comments'][:2]  # Take first 2 comments max
                
            customer_data[file_name] = data
            
        return customer_data
    
    def enhance_process_content(self, response_text):
        """
        Add customer information to process cells in the response text.
        This ensures the total score is visible in the report.
        """
        customer_data = self.extract_customer_data()
        
        if not customer_data:
            return response_text
            
        lines = response_text.split('\n')
        in_table = False
        is_process_table = False
        header_row = []
        process_col_idx = -1
        evidence_col_idx = -1
        
        for i, line in enumerate(lines):
            # Detect start of table
            if line.strip().startswith('|') and line.strip().endswith('|') and not in_table:
                in_table = True
                # Parse header row
                header_cells = [cell.strip().lower() for cell in line.split('|')[1:-1]]
                header_row = header_cells
                
                # Identify if this is the process-evidence table by checking for key columns
                if 'process' in header_cells and 'sighted evidence' in header_cells:
                    is_process_table = True
                    process_col_idx = header_cells.index('process')
                    evidence_col_idx = header_cells.index('sighted evidence')
            
            # Process table rows if we're in a process-evidence table
            elif in_table and is_process_table and line.strip().startswith('|'):
                # Skip separator row
                if all(cell.strip() == '' or all(c in '-:' for c in cell) for cell in line.split('|')[1:-1]):
                    continue
                
                # Get the cells for this row
                cells = line.split('|')
                if len(cells) > process_col_idx + 1 and len(cells) > evidence_col_idx + 1:
                    process_text = cells[process_col_idx + 1].strip()
                    evidence_text = cells[evidence_col_idx + 1].strip()
                    
                    # Match this row to an evidence file
                    matched_file = None
                    for evidence_file in self.evidence_images:
                        if evidence_file.lower() in evidence_text.lower() or any(word in evidence_text.lower() for word in evidence_file.lower().split('.')):
                            matched_file = evidence_file
                            break
                    
                    # Add customer data to process cell if we found a match
                    if matched_file and matched_file in customer_data:
                        data = customer_data[matched_file]
                        
                        # Create formatted customer data
                        formatted_data = []
                        
                        # Start with existing process text if not empty
                        if process_text:
                            formatted_data.append(process_text)
                            formatted_data.append("")  # Add blank line
                        else:
                            # Use file name as process name if cell is empty
                            base_name = os.path.basename(matched_file).split('.')[0]
                            process_name = ' '.join(word.capitalize() for word in base_name.replace('_', ' ').replace('-', ' ').split())
                            formatted_data.append(f"Process: {process_name}")
                            formatted_data.append("")  # Add blank line
                        
                        # Add customer information
                        formatted_data.append("Evidence Details:")
                        
                        if data['company']:
                            formatted_data.append(f"Customer: {data['company']}")
                        if data['date']:
                            formatted_data.append(f"Date: {data['date']}")
                        if data['score']:
                            formatted_data.append(f"Score: {data['score']}")
                        
                        # Add comments if available
                        if data['comments']:
                            formatted_data.append("Comments:")
                            for comment in data['comments']:
                                formatted_data.append(f"- {comment[:100]}...")
                        
                        # Update the process cell with the formatted content
                        cells[process_col_idx + 1] = " " + "\n".join(formatted_data) + " "
                        
                        # Update the line in the response text
                        lines[i] = "|".join(cells)
            
            # Detect end of table
            elif in_table and not line.strip().startswith('|'):
                in_table = False
                is_process_table = False
                header_row = []
                process_col_idx = -1
                evidence_col_idx = -1
        
        return '\n'.join(lines)
    
    def preprocess(self, response_text):
        """Apply all preprocessing steps to the AI response."""
        # Clean table formatting
        cleaned = self.clean_response(response_text)
        
        # Fix heading formats
        cleaned = self.ensure_proper_headings(cleaned)
        
        # Standardize tables
        cleaned = self.standardize_table_format(cleaned)
        
        # Fix checkmark symbols
        cleaned = self.fix_checkmark_symbols(cleaned)
        
        # Enhance process content with customer data
        cleaned = self.enhance_process_content(cleaned)
        
        # Process evidence references and add checkmarks based on scores
        cleaned = self.process_evidence_references(cleaned)
        
        # Extract headers for template filling
        headers = self.extract_headers_from_content(cleaned)
        
        return cleaned
