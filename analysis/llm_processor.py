import streamlit as st
import os
from datetime import datetime
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from dotenv import load_dotenv
import openai
from docx import Document
from openpyxl import  load_workbook
import re
from reportlab.lib.pagesizes import letter
import logging
from .prompts import get_prompt
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

# ------- LLM Integration Module -------
class LLMProcessor:
    """Process extracted text using Language Models with optimized context handling."""
    
    # Define API keys in the backend instead of requiring user input
    OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
    
    # Max tokens for different models
    MODEL_MAX_TOKENS = {
        "gpt-4o": 128000,
        "gpt-4": 8192,
        "gpt-4o-mini": 128000,
    }
    
    # Output tokens to reserve for response
    RESERVED_OUTPUT_TOKENS = {
        "gpt-4o": 4096,
        "gpt-4": 4096,
        "gpt-4o-mini": 4096,
    }
    
    @staticmethod
    def create_audit_prompt(evidence_text, template_structure, auditor_name):
        """Create a highly structured prompt for the LLM to generate a professional audit report."""
        current_date = datetime.now().strftime('%d/%m/%Y')

        prompt = get_prompt(evidence_text=evidence_text, template_structure=template_structure, auditor_name=auditor_name, use_prompt=2) 
        return prompt
        
    @staticmethod
    def estimate_token_count(text):
        """Estimate the number of tokens in a text string.
        Uses a rough approximation of 4 characters per token."""
        return len(text) // 4
    
    @staticmethod
    def get_available_tokens(model):
        """Get the maximum available tokens for input based on model."""
        max_tokens = LLMProcessor.MODEL_MAX_TOKENS.get(model, 4096)
        reserved_tokens = LLMProcessor.RESERVED_OUTPUT_TOKENS.get(model, 2048)
        return max_tokens - reserved_tokens
    
    @staticmethod
    def chunk_evidence(evidence_text, max_tokens, chunk_size=4000):
        """Chunk the evidence text into smaller pieces that fit within token limits.
        Returns a list of evidence chunks."""
        # Split by evidence file markers
        file_pattern = r"--- EVIDENCE FROM (.*?) ---\n\n"
        evidence_parts = re.split(file_pattern, evidence_text)
        
        # Pair filenames with content
        evidence_files = []
        for i in range(1, len(evidence_parts), 2):
            if i < len(evidence_parts) - 1:
                filename = evidence_parts[i]
                content = evidence_parts[i+1]
                evidence_files.append((filename, content))
        
        # Create balanced chunks
        chunks = []
        current_chunk = []
        current_token_count = 0
        
        for filename, content in evidence_files:
            file_header = f"--- EVIDENCE FROM {filename} ---\n\n"
            file_tokens = LLMProcessor.estimate_token_count(content) + LLMProcessor.estimate_token_count(file_header)
            
            # If file is larger than chunk_size, split it further
            if file_tokens > chunk_size:
                # Add file header to current chunk
                if current_token_count + LLMProcessor.estimate_token_count(file_header) <= max_tokens:
                    current_chunk.append(file_header.rstrip())
                    current_token_count += LLMProcessor.estimate_token_count(file_header)
                else:
                    # Start a new chunk
                    if current_chunk:
                        chunks.append("\n\n".join(current_chunk))
                    current_chunk = [file_header.rstrip()]
                    current_token_count = LLMProcessor.estimate_token_count(file_header)
                
                # Split content into paragraphs
                paragraphs = content.split('\n\n')
                for para in paragraphs:
                    para_tokens = LLMProcessor.estimate_token_count(para)
                    
                    if current_token_count + para_tokens <= max_tokens:
                        current_chunk.append(para)
                        current_token_count += para_tokens
                    else:
                        # Add current chunk to chunks and start a new one
                        chunks.append("\n\n".join(current_chunk))
                        current_chunk = [f"(Continued from {filename})\n{para}"]
                        current_token_count = LLMProcessor.estimate_token_count(current_chunk[0])
            else:
                # Add entire file as one unit if it fits in the current chunk
                file_content = f"{file_header}{content}"
                if current_token_count + file_tokens <= max_tokens:
                    current_chunk.append(file_content.rstrip())
                    current_token_count += file_tokens
                else:
                    # Add current chunk to chunks and start a new one with this file
                    chunks.append("\n\n".join(current_chunk))
                    current_chunk = [file_content.rstrip()]
                    current_token_count = file_tokens
        
        # Add the last chunk if it's not empty
        if current_chunk:
            chunks.append("\n\n".join(current_chunk))
        
        return chunks
    
    @staticmethod
    def create_summary_prompt(evidence_chunks, template_structure, *args, **kwargs):
        """Create a prompt to summarize evidence chunks for the final report."""
        current_date = datetime.now().strftime('%d/%m/%Y')
        chunk_summaries = []
        
        for i, chunk in enumerate(evidence_chunks):
            chunk_summaries.append(f"EVIDENCE CHUNK {i+1}:\n{chunk}")
        
        prompt = get_prompt(chunk_summaries=chunk_summaries, template_structure=template_structure, use_prompt=3)
        return prompt
    
    @staticmethod
    def create_final_report_prompt(evidence_summaries, template_structure, auditor_name):
        """Create a prompt to generate the final report based on evidence summaries."""        
        prompt = get_prompt(evidence_summaries=evidence_summaries, template_structure=template_structure, auditor_name=auditor_name, use_prompt=4)
        
        return prompt

    @staticmethod
    def analyze_with_openai(prompt, model="gpt-4o"):
        """Send prompt to OpenAI API and get response."""
        try:
            openai.api_key = LLMProcessor.OPENAI_API_KEY
            
            response = openai.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": "You are an expert ISO auditor assistant specializing in creating detailed audit reports from evidence analysis."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=4096,
                temperature=0.2  # Lower temperature for more factual responses
            )
            
            return response.choices[0].message.content
        except Exception as e:
            st.error(f"Error with OpenAI API: {e}")
            return f"Error: {str(e)}"
    
    
    @staticmethod
    def process_batch_with_openai(evidence_text, template_structure, auditor_name, model="gpt-4o"):
        """Process evidence in batches for OpenAI due to context limitations."""
        try:
            # Calculate available tokens for input
            available_tokens = LLMProcessor.get_available_tokens(model)
            
            # Calculate tokens for template and other parts of the prompt
            template_tokens = LLMProcessor.estimate_token_count(template_structure)
            base_prompt_tokens = LLMProcessor.estimate_token_count(
                LLMProcessor.create_audit_prompt("", template_structure, auditor_name)
            )
            
            # Calculate tokens available for evidence
            evidence_tokens_available = available_tokens - base_prompt_tokens
            
            # If evidence is too large, process in batches
            evidence_tokens = LLMProcessor.estimate_token_count(evidence_text)
            
            if evidence_tokens <= evidence_tokens_available:
                # Evidence fits in one batch
                full_prompt = LLMProcessor.create_audit_prompt(evidence_text, template_structure, auditor_name)
                return LLMProcessor.analyze_with_openai(full_prompt, model)
            else:
                # Need to process in batches
                st.info("Evidence is large - processing in batches for optimal analysis.")
                
                # Chunk the evidence
                evidence_chunks = LLMProcessor.chunk_evidence(
                    evidence_text, 
                    max_tokens=evidence_tokens_available, 
                    chunk_size=evidence_tokens_available//2
                )
                
                # Process each chunk for summary
                chunk_summaries = []
                progress_text = st.empty()
                progress_bar = st.progress(0)
                
                for i, chunk in enumerate(evidence_chunks):
                    progress_text.text(f"Processing evidence chunk {i+1} of {len(evidence_chunks)}...")
                    progress_bar.progress((i) / len(evidence_chunks))
                    
                    # Create a summary prompt for this chunk
                    if i == 0:
                        summary_prompt = LLMProcessor.create_summary_prompt([chunk], template_structure, auditor_name)
                    else:
                        summary_prompt = f"""
                        Continue analyzing this additional evidence chunk:
                        
                        EVIDENCE CHUNK {i+1}:
                        {chunk}
                        
                        Follow the same format as before - identify key findings, issues, and process compliance status.
                        Be specific but concise with bullet points by topic area.
                        """
                    
                    # Get summary for this chunk
                    chunk_summary = LLMProcessor.analyze_with_openai(summary_prompt, model)
                    chunk_summaries.append(f"ANALYSIS OF EVIDENCE CHUNK {i+1}:\n{chunk_summary}")
                
                # Combine summaries and generate final report
                progress_text.text("Generating final comprehensive report...")
                progress_bar.progress(0.9)
                
                combined_summaries = "\n\n".join(chunk_summaries)
                final_prompt = LLMProcessor.create_final_report_prompt(combined_summaries, template_structure, auditor_name)
                
                # Generate final report
                final_report = LLMProcessor.analyze_with_openai(final_prompt, model)
                
                progress_text.text("Report generation complete!")
                progress_bar.progress(1.0)
                
                return final_report
                
        except Exception as e:
            st.error(f"Error processing batches with OpenAI: {e}")
            return f"Error: {str(e)}"
    
    @staticmethod
    def analyze_with_model(prompt, provider="openai", model="gpt-4o", evidence_text="", template_structure="", auditor_name=""):
        """Use the selected AI provider to analyze the prompt."""
        if provider.lower() == "openai":
            # For OpenAI, use batch processing if evidence_text is provided
            if evidence_text and template_structure:
                return LLMProcessor.process_batch_with_openai(evidence_text, template_structure, auditor_name, model)
            else:
                return LLMProcessor.analyze_with_openai(prompt, model)
        elif provider.lower() == "gemini":
            return LLMProcessor.analyze_with_gemini(prompt, model)
        else:
            return f"Error: Unsupported AI provider: {provider}"
        
    @staticmethod
    def extract_corrective_actions(report_content):
        """
        Use LLM to extract concise details and corrective actions from the report.
        
        Args:
            report_content: The full text of the generated audit report
            
        Returns:
            Dictionary with 'details', 'corrective_actions', and 'source_of_issue' keys
        """
        try:
            # Create a targeted prompt for extracting specific information
            prompt = get_prompt(report_content=report_content, use_prompt=5)
            
            # Use a smaller, faster model for this extraction task
            response = LLMProcessor.analyze_with_model(
                prompt=prompt,
                provider="openai", 
                model="gpt-4o-mini"
            )
            
            # Parse the JSON response
            import json
            import re
            
            # Try to extract JSON from the response
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if not isinstance(json_match, type(None)):
                json_str = json_match.group(0)
                result = json.loads(json_str)
                return result
                
            # Fallback parsing if JSON extraction fails
            lines = response.strip().split('\n')
            result = {
                "details": "",
                "corrective_actions": "",
                "source_of_issue": "Internal Audit"  # Default
            }
            
            for line in lines:
                if "details" in line.lower() and ":" in line:
                    result["details"] = line.split(":", 1)[1].strip()
                elif "corrective" in line.lower() and ":" in line:
                    result["corrective_actions"] = line.split(":", 1)[1].strip()
                elif "source" in line.lower() and "issue" in line.lower() and ":" in line:
                    result["source_of_issue"] = line.split(":", 1)[1].strip()
                    
            return result
            
        except Exception as e:
            print(f"Error extracting with LLM: {str(e)}")
            return None

