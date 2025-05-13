from datetime import datetime

def get_prompt(**kwargs):
    current_date = datetime.now().strftime('%d/%m/%Y')

    auditor_name = kwargs.get("auditor_name", "")
    evidence_text = kwargs.get("evidence_text", "")
    template_structure = kwargs.get("template_structure", "")
    chunk_summaries = kwargs.get("chunk_summaries", "")
    evidence_summaries = kwargs.get("evidence_summaries", "")
    report_content = kwargs.get("report_content", "")
    use_prompt = kwargs.get("use_prompt", "")


    audit_prompt_1 = f"""
- **Audit Date:** {current_date}
- **Lead Auditor Designation:** {auditor_name if auditor_name else 'Lead Internal Auditor'}

---

## AUDIT METHODOLOGY REQUIREMENTS

### Evidence Analysis Protocol
1. Conduct systematic examination of ALL provided evidence materials with these specific steps:
   - Extract and catalog all quantitative metrics (satisfaction scores, delivery times, percentages)
   - Identify recurring themes across multiple documents
   - Cross-reference findings between different evidence sources
   - Calculate averages, trends, and pattern deviations where numerical data exists
2. Evaluate compliance status using a 3-step verification approach:
   - Verify document existence and completeness
   - Assess implementation effectiveness through outcomes and metrics
   - Determine conformity with specified ISO requirements
3. Classify and prioritize findings using this framework:
   - Critical: Direct standard violation with significant impact (immediate action required)
   - Major: System deficiency impacting effectiveness (requires prompt attention)
   - Minor: Isolated issue with limited impact (improvement needed)
   - Positive: Exemplary practice exceeding requirements (potential best practice)
4. Document specific objective evidence to support ALL findings, including:
   - Document name and reference number
   - Specific section/page number when available
   - Direct quotes of relevant text where applicable
   - Quantitative data points supporting the conclusion

### Template Completion Standards
1. Complete **all sections** and **all tables** in the audit template without exception
2. Apply evidence-based content where documented evidence exists, including:
   - Direct citation of evidentiary documents with specific references
   - Quantitative metrics with proper statistical context (averages, trends)
   - Cross-referenced information from multiple sources when available
3. Where evidence is insufficient or not available:
   - Apply professional judgment based on ISO standards requirements
   - Clearly denote assumptions with appropriate terminology (e.g., "Based on ISO 9001:2015 clause 8.5.1 requirements..." or "Professional judgment applied due to limited available evidence")
4. Maintain formal, objective audit language throughout the report with:
   - Precise technical terminology aligned with ISO standards
   - Evidence-based assertions rather than opinions
   - Active voice in findings and recommendations

### Compliance Assessment Protocol
1. For each process area assessment:
   - Apply exactly one classification per process (OK/OFI/NC/NA)
   - Utilize "✓" as the standard compliance indicator
   - Support all classifications with minimum 3 specific evidence references or professional rationale
   - Provide substantive comments that explain the basis for each determination
   - Include specific clause references from applicable standards

---

## PROFESSIONAL PRESENTATION STANDARDS

### Document Structure Requirements
- Adhere precisely to the template structure, maintaining all section headers exactly as provided
- Implement proper markdown formatting:
  - Level 1 headings (`#`) for primary sections
  - Level 2 headings (`##`) for subsections where appropriate
  - Properly aligned tables with consistent column structure
- Begin with a concise Executive Summary addressing:
  - Overall conformity status
  - Top 3 strengths identified
  - Top 3 improvement opportunities
  - Critical action recommendations
- Conclude with Final Assessment and Recommendations using the SMART format:
  - Specific - Precisely what needs to be done
  - Measurable - How success will be measured
  - Achievable - Realistic within organizational context
  - Relevant - Aligned with ISO requirements
  - Time-bound - Recommended implementation timeframe

### Professional Terminology
- Utilize ISO-standard terminology consistently throughout the report
- Apply objective, evidence-based language in all assessments
- Maintain formal business communication standards appropriate for executive review
- Use clear, concise language suited for management decision-making

---

## SECTION COMPLETION SPECIFICATIONS

### Core Documentation Sections
- **AUDIT DATE:** Utilize the specified date: {current_date}
- **AUDITOR:** Enter designated auditor: {auditor_name if auditor_name else 'Lead Internal Auditor'}
- **AUDIT ADDRESS:** Extract from evidence or apply appropriate organizational location with full address details


- **RISK DETAILS:** Provide specific evidence-based details for each identified risk, including:
  - Likelihood (High/Medium/Low)
  - Impact (High/Medium/Low)
  - Risk rating (Critical/Significant/Moderate/Minor)
  - Evidence supporting risk assessment

### Compliance Assessment Tables
Apply the following assessment classifications with precision:
- **OK:** Full compliance evidenced (Conformity) - supported by at least 2 specific pieces of evidence
- **OFI:** Opportunity For Improvement identified (Potential Enhancement) - supported by specific observation and standard requirement
- **NC:** Nonconformity detected (Gap requiring correction) - supported by specific evidence and standard clause violation
- **NA:** Not Applicable to this audit scope - supported by scope justification

### Detailed Process Table Format:
```
| PROCESS | SIGHTED EVIDENCE | OK | OFI | NC | NA | ADDITIONAL COMMENTS |
|---------|------------------|----|----|----|----|---------------------|
| [Process Name] | [Document name, ref#, version] [Specific observation] [Interview/demonstration details] | ✓ |  |  |  | [Minimum 2-3 sentences elaborating on evidence and conclusion, with specific standard clause references where applicable] |
```

### Customer Feedback Analysis Section
When customer feedback evidence is provided:
- Calculate overall satisfaction score average across all forms
- Identify highest and lowest scoring categories
- Extract recurring themes from comments sections
- Recommend specific improvements based on feedback patterns
- Cross-reference with quality objectives and performance

### Verification and Follow-up Sections
Provide comprehensive information for all verification elements:
- **NONCONFORMANCES:** Document all identified nonconformities with:
  - Standard clause reference
  - Specific evidence details
  - Risk level classification (Major/Minor)
- **CORRECTIVE ACTIONS:** Recommend appropriate corrective measures with:
  - Root cause analysis requirement
  - Specific action steps
  - Implementation timeframes
  - Verification method
- **OPPORTUNITIES FOR IMPROVEMENT:** Detail all identified enhancement opportunities using format:
  - Current state (As evidenced by...)
  - Desired state (To achieve optimal compliance with [standard/clause]...)
  - Gap analysis (The difference represents...)
  - Specific recommendation (Implementation of...)
- **PREVIOUS AUDIT VERIFICATION:** Document review of previous audit findings with:
  - Confirmation of corrective action implementation
  - Effectiveness assessment with evidence
  - Status classification (Closed/Open/In Progress)
- **RISK MITIGATION VERIFICATION:** Assess effectiveness of existing risk controls with:
  - Evidence of control implementation
  - Measurement of control effectiveness
  - Residual risk assessment
- **COMPETENCY VERIFICATION:** Evaluate personnel qualification evidence including:
  - Training records review
  - Demonstration of competence examples
  - Certification validity verification

### Conclusion Section
- **AUDIT REPORT FINAL COMMENTS:** Provide summative assessment including:
  - Overall management system effectiveness rating
  - Key strengths with supporting evidence
  - Priority improvement areas with rationale
  - Strategic recommendations aligned with organizational objectives
  - Specific next steps with responsible parties and timeframes

---

## REPORT FINALIZATION

Conclude the report with formal signature block:

```
{auditor_name if auditor_name else 'Lead Internal Auditor'}  
Lead Internal Auditor  
{current_date}
```

---

## CRITICAL QUALITY REQUIREMENTS

The final audit report must:
1. Provide comprehensive coverage of ALL evidence provided
2. Contain specific, evidence-based findings for every process area
3. Include quantitative analysis where numerical data exists
4. Present trend analysis across multiple evidence documents
5. Prioritize findings by significance and impact
6. Offer actionable recommendations with implementation guidance
7. Maintain professional audit terminology throughout
8. Distinguish clearly between factual findings and professional judgments
9. Present information in a format suitable for executive decision-making
10. Meet or exceed ISO 19011 Guidelines for auditing management systems

Generate ONLY the completed audit report without any explanatory comments or descriptions of your process.
            """
    audit_prompt_2 = f"""
# ENHANCED ISO INTERNAL AUDIT REPORT GENERATOR

## AUDITOR ROLE SPECIFICATION

You are a certified Lead ISO Internal Auditor specializing in producing formatted audit reports that exactly match the provided template structure.

Your task is to generate a comprehensive, evidence-based Internal Audit Report through systematic analysis of provided documentation. You MUST follow the exact formatting structure shown in the template, including all bullet points, indentation, and section spacing.

---

## INPUT PARAMETERS

- **Evidentiary Documentation:**
```
{evidence_text}
```

- **Audit Template Framework:**
```
{template_structure}
```

- **Audit Date:** {current_date}
- **Lead Auditor Designation:** {auditor_name if auditor_name else 'Lead Internal Auditor'}

---

## CRITICAL AUDIT SECTION REQUIREMENTS 

ALL SECTIONS BELOW MUST BE INCLUDED IN YOUR FINAL REPORT. OMITTING ANY SECTION WILL RESULT IN AN INCOMPLETE AUDIT REPORT.

You MUST provide complete and appropriate content for ALL of the following sections. Remember AUDIT CRITERIA & RISKS AND CAUSES should be tailored to the specific audit focus and not copied verbatim. EACH SECTION REQUIRES AT LEAST 3 SUBSTANTIAL POINTS and must maintain the exact format. Do not leave any section empty:

1. **AUDIT TITLE:** "Internal Audit - Customer Feedback Process"
2. **AUDIT DATE:** {current_date}
3. **AUDITOR:** {auditor_name if auditor_name else 'Lead Internal Auditor'} 
4. **AUDIT ADDRESS:** "Remote internal audit"
5. **AUDIT SCOPE:** "This internal audit applies to the implementation of the organisation's Customer Feedback Process at the Osborne Park location"
6. **AUDIT CRITERIA:** "ISO 9001:2015:

-   4.2 The needs and expectations of interested parties

-   5.1.2 Customer focus

-   9.1.2 Customer satisfaction

Customer Feedback Process PAPROC9.0"
7. **AUDIT PLANNING:** "Internal Audit Schedule PAFORM25.0"
8. **RISKS AND CAUSES:** "Lack of adequate customer feedback causing:

-   Customer dissatisfaction

-   Loss of revenue/opportunities

Unawareness of customer perceptions"
9. **MITIGATION STRATEGIES:** "-   Customer Feedback Process in place and implemented to ensure customer feedback occurs and is analysed

-   Quality target in place to ensure regular customer feedback is conducted

-   Management Review Schedule in place and implemented to ensure customer feedback is analysed by management

-   Regular toolbox meetings conducted to ensure customer feedback results are communicated

Internal audits occur to verify the effectiveness and implementation of the Customer Feedback Process"
10. **LEGAL REQUIREMENTS:** "Nil"
11. **AUDIT REPORT FINAL COMMENTS:** Generate a comprehensive final assessment (MINIMUM 200 WORDS) that includes:
    - Overall management system effectiveness rating
    - Key strengths identified with supporting evidence
    - Priority improvement areas with business impact rationale
    - Strategic recommendations aligned with organizational objectives
    - Specific next steps with responsible parties and timeframes
    - Performance trend comparison with previous audits if applicable
    - Closing statement with auditor's professional judgment
    
    THIS SECTION IS MANDATORY AND MUST BE INCLUDED in the final report.
    
    Conclude with:
    ```
    {auditor_name if auditor_name else 'Internal Auditor'}
    Internal Auditor
    {current_date}
    ```

---

## EVIDENCE TABLE REQUIREMENTS

When filling the PROCESS/SIGHTED EVIDENCE table:

1. Do NOT create duplicate entries for the same evidence or company
2. Each company or process should appear ONLY ONCE in the table
3. Use a consistent naming format: "Process: Company X" for all process cells
4. Scoring System:
   - Place a checkmark "✓" in EXACTLY ONE column (OK, OFI, NC, or NA) per row
   - Use these scoring metrics:
     - OK = 20-25/25 (≥80%)
     - OFI = 18-19/25 (72-76%)
     - NC = <18/25 (<72%)
     - NA = no score available
   - IMPORTANT: Place only ONE checkmark per row
5. Provide detailed, evidence-based comments in the ADDITIONAL COMMENTS column
6. For any reference to evidence, use "Evidence: [filename]" format
7. Use clear, concise language to summarize findings

---

## VERIFICATION AND FOLLOW-UP SECTIONS

CRITICAL INSTRUCTION: You can only have findings in EITHER the NONCONFORMANCES section OR the OPPORTUNITIES FOR IMPROVEMENT section based on the evidence table. If you document nonconformances, you must provide corresponding CORRECTIVE ACTIONS FOR NC and leave the OFI sections empty with "Nil". Conversely, if you document opportunities for improvement, you must provide CORRECTIVE ACTIONS FOR OFI and leave the NC sections empty with "Nil".

### Nonconformance Documentation
- **NONCONFORMANCES:** Document all identified nonconformities with the following structure:
  - NC #1: [Standard/Clause Reference] - [Brief Description]
    - Specific Evidence: Clearly describe what was observed that constitutes a nonconformance
    - Risk Classification: Major/Minor with justification
    - Impact Assessment: Describe potential consequences if not addressed
  
  Example format:
  ```
  NC #1: ISO 9001:2015 Clause 9.1.2 - Inadequate Customer Satisfaction Monitoring
  - Evidence: Customer feedback forms for Company X showed collection frequency of once per year against requirement of quarterly
  - Classification: Minor - System is in place but not fully effective
  - Impact: Limited visibility of changing customer requirements throughout the year
  ```

### Corrective Action Requirements
- **CORRECTIVE ACTIONS FOR NC:** For each nonconformance, provide:
  - Root Cause Analysis Requirement: Specific method recommended (5-Why, Fishbone, etc.)
  - Immediate Containment Actions: What should be done right away
  - Long-term Corrective Measures: Systemic changes needed
  - Implementation Timeframe: Recommended completion dates
  - Verification Method: How effectiveness will be verified
  
  Example format:
  ```
  For NC #1:
  - Root Cause Analysis: Conduct 5-Why analysis to determine reason for infrequent feedback collection
  - Containment: Immediate launch of feedback campaign for all active customers
  - Corrective Action: Implement automated quarterly feedback reminder system with escalation process
  - Timeframe: Implementation within 60 days of audit report
  - Verification: Review of feedback data collection at next internal audit
  ```

### Improvement Opportunity Documentation
- **OPPORTUNITIES FOR IMPROVEMENT:** Detail all identified enhancement opportunities using this structure:
  - OFI #1: [Standard/Clause Reference] - [Brief Description]
    - Current State: "As evidenced by [specific observation]..."
    - Desired State: "To achieve optimal compliance with [standard/clause]..."
    - Gap Analysis: "The difference represents [potential improvement area]..."
    - Specific Recommendation: "Implementation of [detailed suggestion]..."
  
  Example format:
  ```
  OFI #1: ISO 9001:2015 Clause 9.1.2 - Customer Feedback Analysis Enhancement
  - Current State: As evidenced by customer feedback records, basic satisfaction metrics are tracked but minimal trend analysis is performed
  - Desired State: To achieve optimal compliance with clause 9.1.2, comprehensive analysis including correlation of feedback to business performance is recommended
  - Gap Analysis: The difference represents missed opportunities to leverage customer insights for strategic decision-making
  - Specific Recommendation: Implementation of quarterly customer feedback trend analysis with integration into management review agenda
  ```

### Improvement Action Plans
- **CORRECTIVE ACTIONS FOR OFI:** For each improvement opportunity, provide:
  - Implementation Approach: Suggested methodology for improvement
  - Resource Requirements: Personnel, tools, or systems needed
  - Expected Benefits: Quantifiable improvements anticipated
  - Timeline: Suggested implementation schedule
  - Success Metrics: How to measure improvement effectiveness
  
  Example format:
  ```
  For OFI #1:
  - Implementation Approach: Develop structured analysis template with trend visualization
  - Resources: Train quality team on advanced analysis techniques
  - Benefits: Enhanced ability to predict customer needs and prevent dissatisfaction
  - Timeline: Development within 90 days, pilot implementation in Q3
  - Success Metrics: Correlation between identified trends and successful product/service improvements
  ```

### Previous Audit Follow-up
- **WERE PREVIOUS AUDIT RESULTS REVIEWED:** Indicate with "Yes" or "No" if previous audit results were considered, with explanation. If none, write "Nil."

- **GIVE DETAILS OF PREVIOUS AUDIT RESULTS:** Provide a structured summary using:
  - Previous Finding #1: Brief description of finding
  - Original Corrective Action: What was proposed
  - Current Status: Implemented/Partially Implemented/Not Implemented
  - Evidence Reviewed: Documents/processes assessed
  
  If none, write "Nil."

- **WERE PREVIOUSLY IDENTIFIED NONCONFORMANCES OR OPPORTUNITIES FOR IMPROVEMENTS VERIFIED AS CORRECTED, AND WERE CORRECTIVE ACTIONS IMPLEMENTED EFFECTIVE?**
  - Provide a definitive "Yes," "Partially," or "No" with explanation
  - If yes, include evidence of effectiveness
  - If partial or no, explain what remains outstanding
  
  If none, write "Nil."

### Risk Management Evaluation
- **GIVE DETAILS:** Provide specific evidence of risk management effectiveness:
  ```
  Risk mitigation strategies were verified and found to be effective through:
  - Review of [specific risk control documents]
  - Analysis of [performance metrics showing risk reduction]
  - Interviews with [relevant personnel] confirming implementation
  ```

- **WERE RISK MITIGATION STRATEGIES, VERIFIED AS IMPLEMENTED AND EFFECTIVE?** 
  - Answer "Yes" or "No" with supporting evidence
  - If "Yes," provide specific examples of effective implementation
  - If "No," explain deficiencies observed

- **DO THE MITIGATION STRATEGIES RELATED TO THIS PROCESS NEED TO BE REASSESSED FROM NONCONFORMANCES OR OPPORTUNITIES FOR IMPROVEMENT RAISED?**
  - If "Yes," explain why and suggest reassessment approach
  - If "No," justify with reference to audit findings

### Competency Assessment
- **HAVE THE PERSONNEL BEEN VERIFIED AS COMPETENT AS A RESULT OF THIS AUDIT?**
  - Provide "Yes" or "No" with specific evidence:
  - Training records verification details
  - Observed demonstration of required skills
  - Interview results confirming knowledge

### Conclusion Section (MANDATORY)
- **AUDIT REPORT FINAL COMMENTS:** (THIS SECTION MUST BE INCLUDED AND CANNOT BE OMITTED) Provide comprehensive summative assessment (minimum 200 words) including:
  - Overall management system effectiveness rating with scale (e.g., Effective/Partially Effective/Ineffective)
  - Key strengths identified with supporting evidence
  - Priority improvement areas with business impact rationale
  - Strategic recommendations aligned with organizational objectives
  - Specific next steps with responsible parties and timeframes
  - Performance trend comparison with previous audits if applicable
  - Closing statement with auditor's professional judgment
  
  CRITICAL: This section MUST be included in your final report. It is the most important summary section and cannot be omitted under any circumstances. The audit is incomplete without this section.
  
  Conclude with:
  ```
  {auditor_name if auditor_name else 'Internal Auditor'}
  Internal Auditor
  {current_date}
  ```

---

## CRITICAL FORMATTING INSTRUCTIONS

1. BULLET POINT FORMAT: Maintain EXACT same formatting:
   - Use hyphen "-" for bullet points (NOT bullet symbol "•")
   - Keep exact indentation levels as in template
   - Preserve double line breaks between bullet points
   - Do not add or remove bullets from static sections

2. TABLE FORMAT: Tables must maintain their exact column structure and headers

3. DO NOT USE ESCAPE SEQUENCES: When writing line breaks, don't use "\\n" - use actual line breaks

4. NUMERICAL CONSISTENCY: Ensure that if you document 3 nonconformances or OFIs, you must provide 3 corresponding corrective actions

5. EVIDENCE-BASED WRITING: Every finding must be directly linked to specific evidence observed during the audit

---

## DUPLICATE PREVENTION

1. NEVER repeat the same evidence data multiple times
2. Each company or evidence item should appear exactly ONCE
3. Make each process entry unique and clear
4. Do not create empty or "[EMPTY]" rows
5. Ensure each finding references unique evidence

## FINAL REPORT REQUIREMENTS

CRITICAL: Your completed report MUST include ALL sections from this template. Each section is mandatory, especially:

1. All initial audit sections (AUDIT TITLE through LEGAL REQUIREMENTS)
2. The completed PROCESS/SIGHTED EVIDENCE table
3. All verification sections (NONCONFORMANCES through COMPETENCY VERIFICATION)
4. The AUDIT REPORT FINAL COMMENTS section (minimum 200 words)

The AUDIT REPORT FINAL COMMENTS section is a particularly crucial component that must not be omitted. If the model attempts to skip any section, especially the final comments, the report will be considered incomplete and unacceptable.

---

Generate ONLY the completed audit report without any explanatory comments or descriptions of your process. The report must stand as a professional document that could be presented to management without further editing.

I REQUIRE ALL SECTIONS TO BE COMPLETED, ESPECIALLY THE AUDIT REPORT FINAL COMMENTS SECTION. UNDER NO CIRCUMSTANCES SHOULD THIS SECTION BE OMITTED.
"""
  
    summary_prompt_3 = f"""
You are a certified ISO Internal Auditor tasked with creating a professional audit report.

First, analyze these evidence chunks I'm providing from multiple documents:

{chunk_summaries[0] if chunk_summaries else "No evidence provided."}

Additional evidence chunks will be provided separately due to context limitations.

Your final task (after reviewing all evidence) will be to generate a complete audit report following this template structure:

{template_structure}

For now, just identify and note:
1. Key compliance findings from this evidence chunk
2. Any potential issues or gaps identified
3. Any processes mentioned and their compliance status

Format your response as bullet points organized by topic area. Be specific but concise.
"""
  
    final_report_prompt = f"""
You are a certified ISO Internal Auditor.

Based on the evidence analyses I've provided, create a complete, formal Internal Audit Report following the provided template structure.

EVIDENCE ANALYSES:
{evidence_summaries}

TEMPLATE STRUCTURE:
{template_structure}

AUDIT DETAILS:
- Audit Date: {current_date}
- Auditor Name: {auditor_name if auditor_name else 'Internal Auditor'}

REQUIREMENTS:
1. Follow the exact template structure provided
2. Use Markdown formatting with proper headers (#, ##, ###) and table formatting
3. Populate all sections of the template - no sections should be left blank
4. For any areas where evidence is insufficient, make reasonable professional assumptions and note this fact
5. In tables, use "✓" for marking appropriate columns (OK, OFI, NC, NA)
6. Maintain a professional, objective, clear and concise tone throughout
7. Begin with an Executive Summary of key findings
8. End with Recommendations and Final Comments

Format your response as a complete audit report ready for delivery.
    """

    correlative_prompt = f"""
Extract the following information from this internal audit report:

1. Details of issues identified (in 1-2 concise sentences)
2. Appropriate corrective actions (in 1-2 concise sentences)
3. Source of issue (one of: Internal Audit, External Audit, Customer Complaint, Management Review, Process Monitoring, Employee Suggestion)

Format the output as JSON with keys 'details', 'corrective_actions', and 'source_of_issue'.

REPORT CONTENT:
{report_content[:4000]}  # Limit to first 4000 chars for token efficiency
"""
    
    # return prompots
    if use_prompt == 1:
        return audit_prompt_1
    elif use_prompt == 2:
        return audit_prompt_2
    elif use_prompt == 3:
        return summary_prompt_3
    elif use_prompt == 4:
        return final_report_prompt
    elif use_prompt == 5:
        return correlative_prompt
    else:
        raise AssertionError("pass parameter of type dict ->{'use_prompt':val[1?2?3?]}")
    
  

