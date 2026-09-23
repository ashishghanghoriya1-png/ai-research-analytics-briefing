import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def create_element(name):
    return OxmlElement(name)

def set_cell_background(cell, fill_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout_box(doc, title, content, border_color="4F46E5", bg_color="F8FAFC"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border only
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="36" w:space="0" w:color="{border_color}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(title)
    run_t.font.name = 'Calibri'
    run_t.font.size = Pt(11)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(15, 23, 42)
    
    p_c = cell.add_paragraph()
    p_c.paragraph_format.space_before = Pt(0)
    p_c.paragraph_format.space_after = Pt(0)
    run_c = p_c.add_run(content)
    run_c.font.name = 'Calibri'
    run_c.font.size = Pt(10.5)
    run_c.font.color.rgb = RGBColor(51, 65, 85)
    
    # Add spacing after table
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(0)
    p_space.paragraph_format.space_after = Pt(6)

def generate_speaker_notes():
    doc = docx.Document()
    
    # Set page margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Color Palette Constants
    COLOR_PRIMARY = RGBColor(15, 23, 42)      # Slate 900
    COLOR_SECONDARY = RGBColor(79, 70, 229)   # Indigo 600
    COLOR_MUTED = RGBColor(100, 116, 139)     # Slate 500
    COLOR_DARK = RGBColor(30, 41, 59)         # Slate 800
    
    # =========================================================================
    # DOCUMENT COVER / HEADER TITLE
    # =========================================================================
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("Executive Speaker Notes & Presentation Guide")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(26)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("AI Toolchain Architecture for Strategic Research & Analytics | Executive Briefing")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(13)
    run_sub.font.italic = True
    run_sub.font.color.rgb = COLOR_SECONDARY
    
    # Overview Metadata Table
    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    
    headers = [("Presenter Role", "Lead AI Research Architect & Analytics Director"), ("Target Audience", "Executive Leadership & Board"), ("Total Duration", "12 - 15 Minutes"), ("Key Toolchain", "Qwen 2.5, DeepSeek-R1, TabFM, AGY CLI")]
    for j, (h_title, h_val) in enumerate(headers):
        cell = meta_table.cell(0, j)
        cell.width = Inches(1.62)
        set_cell_background(cell, "F1F5F9")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(h_title.upper())
        r1.font.name = 'Calibri'
        r1.font.size = Pt(8.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_MUTED
        
        p2 = cell.paragraphs[1] if len(cell.paragraphs) > 1 else cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(h_val)
        r2.font.name = 'Calibri'
        r2.font.size = Pt(10)
        r2.font.bold = True
        r2.font.color.rgb = COLOR_PRIMARY

    # Spacing
    p_s = doc.add_paragraph()
    p_s.paragraph_format.space_after = Pt(12)

    # Executive Overview Callout
    add_callout_box(
        doc,
        "📌 EXECUTIVE BRIEFING OBJECTIVE",
        "This briefing demonstrates how our research and analytics division has shifted from manual line-by-line qualitative coding to an automated, open-weight AI toolchain (Qwen 2.5 72B + DeepSeek-R1 + TabFM + Streamlit + AGY CLI). The objective is to reassure executive leadership regarding zero cloud infrastructure costs, 100% data privacy, high econometric precision, and immediate policy-level actionability.",
        border_color="4F46E5",
        bg_color="EEF2FF"
    )

    doc.add_heading("Slide-by-Slide Executive Delivery Guide", level=1)

    slides_data = [
        {
            "num": 1,
            "title": "Engineering Modern Research & Analytics with Open AI Toolchains",
            "time": "2 Minutes",
            "anchor": "Executive Hero Bento Card showing the 6 integrated tool badges (Qwen 2.5, DeepSeek-R1, TabFM, Streamlit, Figma, Terminal & AGY CLI).",
            "script": (
                "Good morning, members of the executive team. Today, I am excited to present our updated AI-driven Research & Analytics pipeline.\n\n"
                "Historically, analyzing thousands of qualitative observer notes across Indian public school interventions required weeks of manual coding, leading to reporting lags and subjective variance. Today, we operate a closed-loop, multi-stage open AI stack that reduces qualitative processing time by 70% while elevating predictive statistical fit to 99.1%.\n\n"
                "As you see on screen, our technical architecture orchestrates six specialized open-source tools: Qwen 2.5 72B for qualitative text coding, DeepSeek-R1 for chain-of-thought logic audits, TabFM for econometric modeling, Streamlit for interactive executive portals, Figma for design system specs, and the Antigravity CLI for direct local shell automation. Crucially, this entire stack runs on open-weight models with zero added server subscription costs and complete data privacy."
            ),
            "highlights": [
                "70% reduction in manual reading and qualitative coding effort.",
                "Zero added cloud infrastructure costs via local terminal shell automation.",
                "100% data privacy—no confidential field data sent to external paid API vendors."
            ],
            "qa": [
                ("Q: Why deploy open-weight models locally instead of using proprietary APIs like OpenAI?",
                 "A: Open-weight deployment guarantees 100% data privacy for sensitive state education logs, eliminates unpredictable token billing, and allows fine-grained local execution via terminal automation scripts.")
            ]
        },
        {
            "num": 2,
            "title": "From Observer Field Notes to Tool Telemetry (Multi-Modal Data Intake)",
            "time": "2.5 Minutes",
            "anchor": "Centered Indian government classroom field photo framed with a bold white boundary, flanked by narrative intro on the left and live telemetry metrics cards on the right.",
            "script": (
                "Moving to Slide 2, let's look at where our data originates. On screen in the center is an observation photo from an Indian government classroom framed in our multi-modal intake system.\n\n"
                "Our team ingests 1,268 raw observer field logs written in Hindi, Hinglish, and English across 411 observed classrooms. In traditional research workflows, handling mixed-language handwritten or typed notes creates massive bottlenecks.\n\n"
                "Our pipeline ingests these qualitative narratives directly into local terminal execution scripts. Notice the telemetry metrics on the right: we process multi-lingual transcripts in real-time, feeding qualitative sentiment vectors straight into our predictive models."
            ),
            "highlights": [
                "Sample Size: 1,268 field logs across 411 government school classrooms.",
                "Native multi-lingual NLP handling Hindi, English, and regional Hinglish idioms.",
                "Bold centered visual framing showcasing real-time classroom telemetry intake."
            ],
            "qa": [
                ("Q: How does the system handle informal or dialectal Hindi/English notes?",
                 "A: Qwen 2.5 72B features native multi-lingual tokenization fine-tuned on regional syntax, preserving the exact qualitative meaning of observer comments without requiring pre-translation.")
            ]
        },
        {
            "num": 3,
            "title": "Qualitative & Reasoning Core (01. Qwen 2.5 72B & 02. DeepSeek-R1)",
            "time": "2.5 Minutes",
            "anchor": "Side-by-side technical deep-dive cards comparing Qwen 2.5 (Stage 01 Qualitative LLM) and DeepSeek-R1 (Stage 02 Reasoning Engine).",
            "script": (
                "On Slide 3, we highlight the core analytical engines that power Stages 1 and 2.\n\n"
                "In Stage 1, Qwen 2.5 72B acts as our qualitative copilot. It automatically categorizes 1,268 raw observation passages into 14 standardized policy codebooks, achieving a 12.4x velocity gain over human coding.\n\n"
                "In Stage 2, DeepSeek-R1 serves as our chain-of-thought logic auditor. One major risk with standard LLMs is hallucination. DeepSeek-R1 executes multi-step reasoning to verify Qwen's qualitative themes against raw field transcript lines before any data reaches executive reports. It flags contradictions in 42 milliseconds with 99.4% audit verification.\n\n"
                "Furthermore, our research agent employs 5 specialized methodologies: Dual-Agent Peer Auditing (Qwen ➔ DeepSeek), Qual-to-Quant Vector Fusion (TabFM feature embedding), Native Code-Switching Parsing (Hindi/Hinglish/English), RAG Line-Item Citation Traceability, and Counterfactual What-If Policy Simulation."
            ),
            "highlights": [
                "Qwen 2.5 delivers a 12.4x velocity gain in qualitative taxonomy extraction.",
                "DeepSeek-R1 chain-of-thought (CoT) logic eliminates hallucinations.",
                "42ms reasoning latency with 99.4% verified audit accuracy.",
                "5 Core Research Agent Methods: Peer Audit, Vector Fusion, Code-Switching, RAG Citations, & Counterfactual Simulation."
            ],
            "qa": [
                ("Q: How do we verify that Qwen is not misinterpreting field observations?",
                 "A: DeepSeek-R1 performs automated cross-examination, auditing every extracted theme back to exact source text line numbers before confirming the codebook.")
            ]
        },
        {
            "num": 4,
            "title": "Predictive Econometrics & Local CLI Execution (03. TabFM & 04. Terminal & AGY CLI)",
            "time": "2.5 Minutes",
            "anchor": "Dual infrastructure cards detailing Stage 03 TabFM Econometric Engine and Stage 04 Terminal & Antigravity CLI Execution.",
            "script": (
                "Slide 4 brings us to Stage 3 and Stage 4—the predictive and operational core of our system.\n\n"
                "TabFM is a specialized Tabular Foundation Model. It takes Qwen's qualitative code frequencies—such as teacher kit availability or prep friction—and converts them into quantitative feature vectors. TabFM then runs zero-shot econometric regression (REG context) to predict student attendance and intervention adoption with an extraordinary 99.1% statistical fit.\n\n"
                "In Stage 4, all of this is executed via Terminal & Antigravity CLI. Crucially, our system enforces 100% deterministic pipeline reproducibility: fixed random seeds, versioned open-weight model snapshots, and containerized local CLI shell scripts mean any state auditor or research colleague can re-run the exact terminal commands and reproduce identical findings every single time."
            ),
            "highlights": [
                "TabFM achieves 99.1% zero-shot REG (econometric regression) fit on qualitative-quantitative datasets.",
                "Deterministic Pipeline Reproducibility via seed-locked terminal scripts and line-numbered audit logs.",
                "Terminal & Antigravity CLI execution provides 100% private, local shell automation."
            ],
            "qa": [
                ("Q: How does the pipeline ensure scientific and audit reproducibility?",
                 "A: By executing seed-locked local scripts via Antigravity CLI, every step—from qualitative text parsing to TabFM econometric REG estimation—is 100% deterministic and reproducible across audit cycles."),
                ("Q: What makes TabFM's REG model superior to standard regression?",
                 "A: TabFM's Zero-Shot Tabular Regression (REG) is pre-trained to model sparse, non-linear relationships across small sample sizes (N=60) where standard OLS regression fails due to missing survey variables.")
            ]
        },
        {
            "num": 5,
            "title": "Interactive Analytics Portals & Design Systems (05. Streamlit & 06. Figma)",
            "time": "2.5 Minutes",
            "anchor": "Bento card showcasing the live Streamlit Cloud Portals (EDS & CRO buttons + PDF download) alongside Figma UI design token specs.",
            "script": (
                "Slide 5 illustrates how these insights are delivered to leadership and program managers.\n\n"
                "We deploy production-grade interactive analytics portals via Streamlit Cloud. Executives do not need to read raw spreadsheets; they can open live web portals to slice qualitative observer notes against quantitative district filters in real-time.\n\n"
                "As shown on screen, we have two live production portals deployed:\n"
                "1. The EDS Streamlit Dashboard at cm-rise-eds-dashboard.streamlit.app\n"
                "2. The CRO Streamlit Dashboard at peepul-mp-cpd-cro-dashboard.streamlit.app\n\n"
                "Additionally, users can download the official 9-page research publication, EDS_Report_TabFM.pdf, directly from the portal interface. All interfaces are styled using Figma UI design tokens for high-craft readability."
            ),
            "highlights": [
                "Live Production Portals: EDS Dashboard & CRO Dashboard deployed on Streamlit Cloud.",
                "Direct one-click export for clean SPSS syntax files and R dataframes.",
                "Official publication download: EDS_Report_TabFM.pdf (Qualitative Field Insights)."
            ],
            "qa": [
                ("Q: Can non-technical policy leads export data for formal academic or government publication?",
                 "A: Yes, the Streamlit portal includes one-click export buttons for SPSS syntax, R dataframes, and clean CSV tables.")
            ]
        },
        {
            "num": 6,
            "title": "Paradigm Shift Comparison & Empirical Case Study Proofs",
            "time": "3 Minutes",
            "anchor": "3-Stage Paradigm Shift Table (Field Insights -> LLMs -> TabFM) and dual Case Study Proof Cards (EDS N=60 & CRO 411 Classrooms).",
            "script": (
                "Finally, Slide 6 synthesizes our technical evolution and presents two concrete policy case studies.\n\n"
                "Looking at the top comparison table, we see our paradigm shift: Stage 1 field notes are transformed by Stage 2 LLMs into qualitative codes, which Stage 3 TabFM turns into 99.1% accurate predictive forecasts.\n\n"
                "Look at the EDS Case Study on the left: analyzing 60 observer notes, Qwen and TabFM revealed that 31.6% of teachers lacked physical reference kits. Leadership responded immediately by mandating physical kit pairing for all future sessions.\n\n"
                "On the right, the CRO Study analyzed 1,268 field notes across 411 classrooms, proving that low lesson alignment (3.6%) was driven by preparation friction rather than teacher resistance. Both case cards feature direct links to launch the live Streamlit apps and download the official EDS TabFM PDF report."
            ),
            "highlights": [
                "EDS Study: Identified 31.6% material shortage, driving immediate kit pairing mandate.",
                "CRO Study: Discovered 3.6% alignment was caused by prep friction across 411 classrooms.",
                "Paradigm shift from manual reporting to dual-engine qual+quant predictive forecasting."
            ],
            "qa": [
                ("Q: How quickly can leadership turn these insights into policy decisions?",
                 "A: Insights that previously took 4-6 weeks are now generated in real-time, allowing policy responses (like kit distribution mandates) within days of field data collection.")
            ]
        }
    ]

    for slide in slides_data:
        doc.add_heading(f"Slide {slide['num']}: {slide['title']}", level=2)
        
        # Meta info
        p_m = doc.add_paragraph()
        p_m.paragraph_format.space_before = Pt(2)
        p_m.paragraph_format.space_after = Pt(4)
        r_t = p_m.add_run(f"⏱️ Allocated Delivery Time: {slide['time']}  |  🎯 Focus: Executive Decision Support")
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(10)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_SECONDARY

        # Visual Anchor
        add_callout_box(
            doc,
            "👁️ VISUAL ANCHOR (WHAT IS ON SCREEN)",
            slide['anchor'],
            border_color="64748B",
            bg_color="F8FAFC"
        )

        # Spoken Script
        doc.add_heading("🗣️ Spoken Script (Word-for-Word Talking Points)", level=3)
        p_script = doc.add_paragraph()
        p_script.paragraph_format.space_before = Pt(2)
        p_script.paragraph_format.space_after = Pt(8)
        p_script.paragraph_format.line_spacing = 1.15
        r_s = p_script.add_run(slide['script'])
        r_s.font.name = 'Calibri'
        r_s.font.size = Pt(11)
        r_s.font.color.rgb = COLOR_PRIMARY

        # Key Highlights
        doc.add_heading("⭐ Key Points to Emphasize", level=3)
        for h in slide['highlights']:
            p_h = doc.add_paragraph(style='List Bullet')
            p_h.paragraph_format.space_before = Pt(0)
            p_h.paragraph_format.space_after = Pt(2)
            r_h = p_h.add_run(h)
            r_h.font.name = 'Calibri'
            r_h.font.size = Pt(10.5)
            r_h.font.color.rgb = COLOR_DARK

        # Executive Q&A
        doc.add_heading("❓ Anticipated Executive Q&A", level=3)
        for q, a in slide['qa']:
            add_callout_box(
                doc,
                q,
                a,
                border_color="0D9488",
                bg_color="F0FDFA"
            )
            
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # =========================================================================
    # APPENDIX: LIVE PORTALS & OFFICIAL REPORT REFERENCES
    # =========================================================================
    doc.add_heading("Appendix: Live System Resources & Report Files", level=1)
    
    app_table = doc.add_table(rows=4, cols=3)
    app_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    app_table.autofit = False
    
    col_widths = [Inches(1.8), Inches(3.2), Inches(1.5)]
    headers = ["Resource Name", "URL / Local Path", "Access Status"]
    
    hdr_cells = app_table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "0F172A")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        r = p.add_run(title)
        r.font.name = 'Calibri'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    resources = [
        ("EDS Streamlit Dashboard", "https://cm-rise-eds-dashboard.streamlit.app/", "Live Cloud Portal"),
        ("CRO Streamlit Dashboard", "https://peepul-mp-cpd-cro-dashboard.streamlit.app/", "Live Cloud Portal"),
        ("Official TabFM Report PDF", "EDS_Report_TabFM.pdf (9 Pages)", "Local PDF Report")
    ]

    for row_idx, (r_name, r_url, r_status) in enumerate(resources, start=1):
        row_cells = app_table.rows[row_idx].cells
        for c_idx, val in enumerate([r_name, r_url, r_status]):
            row_cells[c_idx].width = col_widths[c_idx]
            bg = "F8FAFC" if row_idx % 2 == 0 else "FFFFFF"
            set_cell_background(row_cells[c_idx], bg)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=100, right=100)
            p = row_cells[c_idx].paragraphs[0]
            r = p.add_run(val)
            r.font.name = 'Calibri'
            r.font.size = Pt(9.5)
            r.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                r.font.bold = True

    # =========================================================================
    # PROGRAM LEADERSHIP & TEAM CREDITS
    # =========================================================================
    doc.add_heading("Program Leadership & Team Credits", level=1)
    
    add_callout_box(
        doc,
        "Architectural Ideation & Program Context: Shilish & Shil",
        "Shilish presented the original concept and program-level code design that laid the groundwork for this AI stack. Shil helped establish and shape the program context and overall analytical framework.",
        border_color="D97706",
        bg_color="FFFBEB"
    )
    
    add_callout_box(
        doc,
        "Strategic Guidance & Program Oversight: Rohan",
        "Rohan provided continuous executive leadership, strategic direction, and project guidance throughout the entire initiative.",
        border_color="2563EB",
        bg_color="EFF6FF"
    )
    
    add_callout_box(
        doc,
        "Development & Lead Engineering: Ashish",
        "Ashish served as the lead developer and engineer, responsible for end-to-end implementation, model orchestration, interactive dashboard development, and cloud deployment.",
        border_color="4F46E5",
        bg_color="EEF2FF"
    )

    out_file = "AI_Research_Analytics_Executive_Speaker_Notes.docx"
    doc.save(out_file)
    print(f"Executive Speaker Notes saved successfully: {os.path.abspath(out_file)}")

if __name__ == "__main__":
    generate_speaker_notes()
