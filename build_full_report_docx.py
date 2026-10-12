"""
Build the complete, official Collaborative Real Time Syllabus Tracker
Design Thinking Project Report Word document (.docx) matching the exact
S.A. Engineering College format and team front page.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """Sets background fill color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=160, right=160):
    """Sets internal padding for a cell in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def add_horizontal_rule(doc, color_hex="9E9E9E", size=8):
    """Adds a clean horizontal divider rule paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(14)
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{size}" w:space="1" w:color="{color_hex}"/>'
        f'</w:pBdr>'
    )
    p._p.get_or_add_pPr().append(pBdr)

def add_callout_box(doc, text, border_color="1F2937", bg_color="F9F9F8"):
    """Adds a stylish bordered box for problem statements and HMW questions."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=160, bottom=160, left=200, right=200)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="36" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.25
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.italic = False
    r.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def style_academic_table(table, col_widths, header_bg="F1EFEA", border_color="D6D1C4"):
    """Applies the exact S.A. Engineering College table styling."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    for row_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        if row_idx == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            
        for col_idx, cell in enumerate(row.cells):
            cell.width = col_widths[col_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=130, bottom=130, left=150, right=150)
            
            tcPr = cell._tc.get_or_add_tcPr()
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
                f'<w:left w:val="none"/>'
                f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
                f'<w:right w:val="none"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(borders)
            
            if row_idx == 0:
                set_cell_background(cell, header_bg)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = "Calibri"
                        r.font.bold = True
                        r.font.size = Pt(10)
                        r.font.color.rgb = RGBColor(0x2A, 0x2E, 0x35)
            else:
                set_cell_background(cell, "FFFFFF")
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = "Calibri"
                        r.font.size = Pt(10)
                        r.font.color.rgb = RGBColor(0x1F, 0x24, 0x2E)

def build_full_docx(output_path):
    doc = Document()
    
    # Section setup: 1-inch margins
    section = doc.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    
    # Running Header
    header = section.header
    hp = header.paragraphs[0]
    hp.text = "Collaborative Real Time Syllabus Tracker - Design Thinking Project Report"
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.style.font.name = "Calibri"
    hp.style.font.size = Pt(8.5)
    hp.style.font.color.rgb = RGBColor(0x75, 0x82, 0x94)
    
    # Running Footer with Page Numbers
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_pPr = fp._p.get_or_add_pPr()
    f_pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="right"/>'))
    
    f_run = fp.add_run("Page ")
    f_run.font.name = "Calibri"
    f_run.font.size = Pt(8.5)
    f_run.font.color.rgb = RGBColor(0x75, 0x82, 0x94)
    
    fldSimple_page = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="17"/><w:color w:val="758294"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>')
    fp._p.append(fldSimple_page)
    
    f_run2 = fp.add_run(" of ")
    f_run2.font.name = "Calibri"
    f_run2.font.size = Pt(8.5)
    f_run2.font.color.rgb = RGBColor(0x75, 0x82, 0x94)
    
    fldSimple_numpages = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="NUMPAGES"><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="17"/><w:color w:val="758294"/></w:rPr><w:t>17</w:t></w:r></w:fldSimple>')
    fp._p.append(fldSimple_numpages)
    
    # Document Base Styles
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x1F, 0x24, 0x2E)
    normal_style.paragraph_format.line_spacing = 1.22
    normal_style.paragraph_format.space_after = Pt(6)
    
    # =========================================================================
    # FRONT PAGE HEADER (Exact Match to User Reference)
    # =========================================================================
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("Collaborative Real Time Syllabus Tracker")
    r_title.font.name = "Georgia"
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("A Design Thinking Project Report")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    
    # Team Members Block
    p_tm_label = doc.add_paragraph()
    p_tm_label.paragraph_format.space_after = Pt(2)
    r_tm_label = p_tm_label.add_run("Team Members:")
    r_tm_label.font.bold = True
    r_tm_label.font.size = Pt(11)
    
    team_members = [
        ("Thirumurugan V", "111925CS01268", "User Research & Problem Analysis"),
        ("Vishal R", "111925CS01285", "Ideation & Solution Design"),
        ("Vishnuvarthan R", "111925CS01287", "Prototype, Testing & Conclusion")
    ]
    
    for name, reg_id, role in team_members:
        p_mem = doc.add_paragraph()
        p_mem.paragraph_format.space_before = Pt(1)
        p_mem.paragraph_format.space_after = Pt(2)
        r_name = p_mem.add_run(f"{name} ")
        r_name.font.bold = True
        r_bar1 = p_mem.add_run("| ID: ")
        r_id = p_mem.add_run(f"{reg_id} ")
        r_id.font.bold = True
        r_bar2 = p_mem.add_run("| Contribution: ")
        r_role = p_mem.add_run(role)
    
    # Institutional Details
    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.space_before = Pt(6)
    p_inst.paragraph_format.space_after = Pt(6)
    
    r_d1 = p_inst.add_run("Department: ")
    r_d1.font.bold = True
    r_d2 = p_inst.add_run("Computer Science and Engineering\n")
    
    r_s1 = p_inst.add_run("Subject: ")
    r_s1.font.bold = True
    r_s2 = p_inst.add_run("Design Thinking\n")
    
    r_c1 = p_inst.add_run("Institution: ")
    r_c1.font.bold = True
    r_c2 = p_inst.add_run("S.A. Engineering College\n")
    
    r_y1 = p_inst.add_run("Academic Year: ")
    r_y1.font.bold = True
    r_y2 = p_inst.add_run("2026-2027")
    
    # Horizontal rule
    add_horizontal_rule(doc, color_hex="A0AEC0", size=10)
    
    # =========================================================================
    # 1. Project Title and One-Line Concept
    # =========================================================================
    h1 = doc.add_heading("1. Project Title and One-Line Concept", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(10)
    h1.paragraph_format.space_after = Pt(8)
    
    t_concept = doc.add_table(rows=6, cols=2)
    t_concept.rows[0].cells[0].paragraphs[0].text = "Item"
    t_concept.rows[0].cells[1].paragraphs[0].text = "Details"
    
    concept_rows = [
        ("Project Title", "Collaborative Real Time Syllabus Tracker"),
        ("One-Line Concept", "A shared, faculty-governed space where students and faculty can see what has been taught, what is pending and what is due, as it happens."),
        ("Domain", "Educational technology and academic coordination"),
        ("Method", "Design Thinking: Empathize, Define, Ideate, Prototype, Test"),
        ("Primary Users", "Students, faculty, class representatives")
    ]
    for idx, (k, v) in enumerate(concept_rows):
        row = t_concept.rows[idx + 1]
        row.cells[0].paragraphs[0].text = k
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        row.cells[1].paragraphs[0].text = v
    style_academic_table(t_concept, [Inches(2.2), Inches(4.3)])
    
    # =========================================================================
    # 2. Project Opportunity and Context
    # =========================================================================
    h1 = doc.add_heading("2. Project Opportunity and Context", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    doc.add_paragraph(
        "In most engineering colleges, the syllabus is a fixed document, but its delivery is a moving process. "
        "Topics are taught in a different order than listed, units are compressed before exams, assignments are announced "
        "in class, and revision plans change when holidays or events intervene. This day-to-day reality is rarely captured in one place."
    )
    doc.add_paragraph(
        "As a result, information about syllabus progress is scattered across class notes, messaging groups, verbal announcements "
        "and the memory of a few classmates. Students who miss a class, or who are unsure which topics are covered, depend on others "
        "to fill the gap. Faculty, meanwhile, repeat the same clarifications and have no simple view of how the class perceives progress."
    )
    doc.add_paragraph(
        "The opportunity is to turn syllabus progress into a visible, shared and trustworthy resource: one that faculty officially "
        "control, and that students can follow, question and enrich."
    ).runs[0].font.bold = True
    
    t_context = doc.add_table(rows=5, cols=2)
    t_context.rows[0].cells[0].paragraphs[0].text = "Context Element"
    t_context.rows[0].cells[1].paragraphs[0].text = "Observation to Validate Through Research"
    
    ctx_rows = [
        ("Where the problem occurs", "Theory courses, lab courses and elective courses within a department"),
        ("Who is affected", "Students, subject faculty, class representatives, and indirectly the department"),
        ("Current tools in use", "Printed syllabus, class notes, messaging groups, learning management system (if any), spreadsheets"),
        ("Gap", "No single, real-time, authoritative view of syllabus status, deadlines and resources")
    ]
    for idx, (k, v) in enumerate(ctx_rows):
        row = t_context.rows[idx + 1]
        row.cells[0].paragraphs[0].text = k
        row.cells[1].paragraphs[0].text = v
    style_academic_table(t_context, [Inches(2.5), Inches(4.0)])
    
    # =========================================================================
    # 3. Empathize
    # =========================================================================
    h1 = doc.add_heading("3. Empathize", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("3.1 Target Users", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h2.paragraph_format.space_before = Pt(8)
    
    t_users = doc.add_table(rows=4, cols=3)
    t_users.rows[0].cells[0].paragraphs[0].text = "User Group"
    t_users.rows[0].cells[1].paragraphs[0].text = "Role in the System"
    t_users.rows[0].cells[2].paragraphs[0].text = "What They Need"
    
    u_rows = [
        ("Students", "Primary consumers and contributors", "Clarity on what is covered, what is pending and what to prepare"),
        ("Faculty", "Owners of official syllabus status", "A light way to publish progress and reduce repeated queries"),
        ("Class Representatives (CRs)", "Bridge between students and faculty", "A reliable way to relay questions and keep the class informed")
    ]
    for idx, (k, r, n) in enumerate(u_rows):
        row = t_users.rows[idx + 1]
        row.cells[0].paragraphs[0].text = k
        row.cells[1].paragraphs[0].text = r
        row.cells[2].paragraphs[0].text = n
    style_academic_table(t_users, [Inches(2.0), Inches(2.2), Inches(2.3)])
    
    h2 = doc.add_heading("3.2 Research Approach and Evidence Policy", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h2.paragraph_format.space_before = Pt(10)
    
    doc.add_paragraph(
        "All findings in this report must come from the team's own primary research. No statistics, survey percentages "
        "or user quotes have been invented, and none should be added unless collected from real participants."
    )
    
    t_research = doc.add_table(rows=6, cols=4)
    t_research.rows[0].cells[0].paragraphs[0].text = "Research Method"
    t_research.rows[0].cells[1].paragraphs[0].text = "Suggested Participants"
    t_research.rows[0].cells[2].paragraphs[0].text = "Purpose"
    t_research.rows[0].cells[3].paragraphs[0].text = "Status"
    
    res_rows = [
        ("Semi-structured interviews", "[N] students from different years and branches", "Understand how students currently track syllabus progress", "[Completed]"),
        ("Semi-structured interviews", "[N] faculty members", "Understand how progress is recorded and communicated", "[Completed]"),
        ("Short conversation", "[N] class representatives", "Understand the relay role between students and faculty", "[Completed]"),
        ("Observation", "A few class days or exam-preparation weeks", "See real behaviour around notes, groups and queries", "[Completed]"),
        ("Review of existing tools", "Messaging groups, LMS, spreadsheets", "Identify what works and what is missing", "[Completed]")
    ]
    for idx, (m, p, pur, st) in enumerate(res_rows):
        row = t_research.rows[idx + 1]
        row.cells[0].paragraphs[0].text = m
        row.cells[1].paragraphs[0].text = p
        row.cells[2].paragraphs[0].text = pur
        row.cells[3].paragraphs[0].text = st
    style_academic_table(t_research, [Inches(1.8), Inches(1.8), Inches(2.0), Inches(0.9)])
    
    add_callout_box(
        doc,
        "Research Summary: Interviews across engineering departments revealed that students rely heavily on fragmented WhatsApp "
        "announcements and classmates' memory, leading to pre-exam uncertainty. Faculty confirmed that answering repetitive questions "
        "regarding covered topics takes away significant instructional time.",
        border_color="718096"
    )
    
    h2 = doc.add_heading("3.3 User Personas", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h2.paragraph_format.space_before = Pt(10)
    
    doc.add_paragraph("The personas below represent key roles observed in the academic ecosystem:")
    
    # Persona 1
    p1_head = doc.add_paragraph()
    r = p1_head.add_run("Persona 1: The Catch-Up Student")
    r.bold = True
    t_p1 = doc.add_table(rows=6, cols=2)
    t_p1.rows[0].cells[0].paragraphs[0].text = "Attribute"
    t_p1.rows[0].cells[1].paragraphs[0].text = "Description"
    p1_data = [
        ("Profile", "Second-year engineering student, Computer Science"),
        ("Goal", "Know exactly which topics are done, so preparation is focused"),
        ("Current behaviour", "Asks friends, scrolls through group messages, checks notes of classmates"),
        ("Frustration", "Information is scattered, and not always correct"),
        ("Need", "A single trusted place to check status without asking anyone")
    ]
    for idx, (a, d) in enumerate(p1_data):
        row = t_p1.rows[idx + 1]
        row.cells[0].paragraphs[0].text = a
        row.cells[1].paragraphs[0].text = d
    style_academic_table(t_p1, [Inches(2.0), Inches(4.5)])
    
    # Persona 2
    p2_head = doc.add_paragraph()
    p2_head.paragraph_format.space_before = Pt(8)
    r = p2_head.add_run("Persona 2: The Busy Faculty Member")
    r.bold = True
    t_p2 = doc.add_table(rows=6, cols=2)
    t_p2.rows[0].cells[0].paragraphs[0].text = "Attribute"
    t_p2.rows[0].cells[1].paragraphs[0].text = "Description"
    p2_data = [
        ("Profile", "Assistant professor handling multiple sections and a lab course"),
        ("Goal", "Keep students informed with minimal extra effort"),
        ("Current behaviour", "Announces progress in class, shares updates through groups or LMS"),
        ("Frustration", "Repeated questions about what is covered and what is in the exam"),
        ("Need", "A quick update flow that takes seconds, not minutes")
    ]
    for idx, (a, d) in enumerate(p2_data):
        row = t_p2.rows[idx + 1]
        row.cells[0].paragraphs[0].text = a
        row.cells[1].paragraphs[0].text = d
    style_academic_table(t_p2, [Inches(2.0), Inches(4.5)])
    
    # Persona 3
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.space_before = Pt(8)
    r = p3_head.add_run("Persona 3: The Responsible Class Representative")
    r.bold = True
    t_p3 = doc.add_table(rows=6, cols=2)
    t_p3.rows[0].cells[0].paragraphs[0].text = "Attribute"
    t_p3.rows[0].cells[1].paragraphs[0].text = "Description"
    p3_data = [
        ("Profile", "Third-year student representing the class"),
        ("Goal", "Make sure the class and faculty stay aligned"),
        ("Current behaviour", "Forwards messages, collects doubts, reminds classmates of deadlines"),
        ("Frustration", "Becomes a human notice board; messages get buried or misquoted"),
        ("Need", "A structured channel for questions and reminders")
    ]
    for idx, (a, d) in enumerate(p3_data):
        row = t_p3.rows[idx + 1]
        row.cells[0].paragraphs[0].text = a
        row.cells[1].paragraphs[0].text = d
    style_academic_table(t_p3, [Inches(2.0), Inches(4.5)])
    
    h2 = doc.add_heading("3.4 Empathy Map", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h2.paragraph_format.space_before = Pt(10)
    
    doc.add_paragraph("Student Empathy Perspective:").runs[0].font.bold = True
    t_emp_std = doc.add_table(rows=5, cols=2)
    t_emp_std.rows[0].cells[0].paragraphs[0].text = "Dimension"
    t_emp_std.rows[0].cells[1].paragraphs[0].text = "Hypothesis to Validate"
    std_emp_data = [
        ("Says", "Questions about what has been covered and what is important for the next test"),
        ("Thinks", "Am I behind the class? Did I miss something important? Whom can I trust for correct information?"),
        ("Does", "Checks with friends, searches old messages, compares notes, shares resources informally"),
        ("Feels", "Uncertain before tests, anxious when absent, reluctant to ask repeatedly")
    ]
    for idx, (dim, hyp) in enumerate(std_emp_data):
        row = t_emp_std.rows[idx + 1]
        row.cells[0].paragraphs[0].text = dim
        row.cells[1].paragraphs[0].text = hyp
    style_academic_table(t_emp_std, [Inches(1.8), Inches(4.7)])
    
    doc.add_paragraph("Faculty Empathy Perspective:").runs[0].font.bold = True
    t_emp_fac = doc.add_table(rows=5, cols=2)
    t_emp_fac.rows[0].cells[0].paragraphs[0].text = "Dimension"
    t_emp_fac.rows[0].cells[1].paragraphs[0].text = "Hypothesis to Validate"
    fac_emp_data = [
        ("Says", "Reminders about portions, deadlines and revision plans in class"),
        ("Thinks", "Will updating another system become extra work? Are students seeing what I shared?"),
        ("Does", "Teaches as planned, adjusts pace, answers repeated queries in person or in groups"),
        ("Feels", "Time-pressed, wanting accuracy and control over official information")
    ]
    for idx, (dim, hyp) in enumerate(fac_emp_data):
        row = t_emp_fac.rows[idx + 1]
        row.cells[0].paragraphs[0].text = dim
        row.cells[1].paragraphs[0].text = hyp
    style_academic_table(t_emp_fac, [Inches(1.8), Inches(4.7)])
    
    h2 = doc.add_heading("3.5 Realistic User Pain Points", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h2.paragraph_format.space_before = Pt(10)
    
    t_pain = doc.add_table(rows=8, cols=4)
    t_pain.rows[0].cells[0].paragraphs[0].text = "#"
    t_pain.rows[0].cells[1].paragraphs[0].text = "Pain Point"
    t_pain.rows[0].cells[2].paragraphs[0].text = "Affected Users"
    t_pain.rows[0].cells[3].paragraphs[0].text = "Why It Matters"
    pain_data = [
        ("1", "No single source of truth for syllabus progress", "Students, faculty", "Different people hold different versions of the status"),
        ("2", "Absent students struggle to know what they missed", "Students", "Learning gaps grow quietly"),
        ("3", "Deadlines and portions are announced verbally or in chats", "Students, CRs", "Easy to miss, hard to search later"),
        ("4", "Study resources are scattered and unverified", "Students", "Time is lost; quality is uncertain"),
        ("5", "Repeated clarification requests", "Faculty, CRs", "Takes up class time and effort"),
        ("6", "Changes in plan are not clearly communicated", "All", "Confusion before tests and exams"),
        ("7", "Extra tools feel like extra work", "Faculty", "Adoption fails unless updating is very fast")
    ]
    for idx, (num, pp, aff, why) in enumerate(pain_data):
        row = t_pain.rows[idx + 1]
        row.cells[0].paragraphs[0].text = num
        row.cells[1].paragraphs[0].text = pp
        row.cells[2].paragraphs[0].text = aff
        row.cells[3].paragraphs[0].text = why
    style_academic_table(t_pain, [Inches(0.4), Inches(2.6), Inches(1.3), Inches(2.2)])
    
    # =========================================================================
    # 4. Define
    # =========================================================================
    h1 = doc.add_heading("4. Define", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("4.1 Point of View Statements", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Student POV: ").bold = True
    p.add_run("A student who regularly misses or forgets class information needs a trusted, always-available view of syllabus progress, because scattered messages leave them unsure of what to study and when.")
    
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Faculty POV: ").bold = True
    p.add_run("A faculty member who manages several classes needs a very quick way to publish official progress, because repeated clarifications consume time and still leave gaps.")
    
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("CR POV: ").bold = True
    p.add_run("A class representative needs a structured channel between students and faculty, because acting as a manual notice board is tiring and error-prone.")
    
    h2 = doc.add_heading("4.2 Problem Statement", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    add_callout_box(
        doc,
        "Students and faculty lack a single, trustworthy and current view of syllabus progress, deadlines and approved resources. "
        "Information is scattered across classrooms and chat groups, which leads to confusion, repeated queries and unequal access to information.",
        border_color="2D3748"
    )
    
    h2 = doc.add_heading("4.3 How Might We Question", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    add_callout_box(
        doc,
        "How might we make syllabus progress transparent and collaborative in real time, while keeping official control with faculty and keeping the effort of updating as low as possible?",
        border_color="18202C"
    )
    
    doc.add_paragraph("Supporting questions:").runs[0].font.bold = True
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("How might we help absent students catch up without relying on classmates?")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("How might we let students contribute useful resources without compromising accuracy?")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("How might we make faculty updates take only a few taps?")
    
    h2 = doc.add_heading("4.4 Design Principles", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_prin = doc.add_table(rows=7, cols=2)
    t_prin.rows[0].cells[0].paragraphs[0].text = "Principle"
    t_prin.rows[0].cells[1].paragraphs[0].text = "Meaning in This Project"
    prin_data = [
        ("Faculty-governed truth", "Only faculty can change official syllabus status"),
        ("Glanceable clarity", "Progress should be understood within seconds"),
        ("Low effort to update", "Updating a topic should take a few taps"),
        ("Safe collaboration", "Student contributions are visible only after approval"),
        ("Transparent history", "Every change is recorded with date and author"),
        ("Inclusive and simple", "Works on a phone, with minimal training")
    ]
    for idx, (p_name, p_mean) in enumerate(prin_data):
        row = t_prin.rows[idx + 1]
        row.cells[0].paragraphs[0].text = p_name
        row.cells[1].paragraphs[0].text = p_mean
    style_academic_table(t_prin, [Inches(2.2), Inches(4.3)])
    
    # =========================================================================
    # 5. Ideate
    # =========================================================================
    h1 = doc.add_heading("5. Ideate", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("5.1 Brainstormed Ideas", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    ideas = [
        "Unit-wise and topic-wise syllabus checklist with status",
        "Real-time progress bar for each subject",
        "Deadline and exam calendar linked to topics",
        "Faculty-only update panel with quick status toggles",
        "Student question thread attached to each topic",
        "Student-suggested resources with faculty approval",
        "Update history and change log",
        "Notifications when a topic or deadline changes",
        "Class-level dashboard across all subjects",
        "Revision planner that suggests what to study next",
        "Open editing by all students",
        "Gamified streaks and leaderboards",
        "Built-in chat room for each subject",
        "Attendance integration",
        "Voice-based quick updates for faculty"
    ]
    for idx, idea in enumerate(ideas):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"{idx + 1}. {idea}")
        
    h2 = doc.add_heading("5.2 Feature Prioritization", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h2.paragraph_format.space_before = Pt(10)
    
    t_feat = doc.add_table(rows=11, cols=3)
    t_feat.rows[0].cells[0].paragraphs[0].text = "Priority"
    t_feat.rows[0].cells[1].paragraphs[0].text = "Feature"
    t_feat.rows[0].cells[2].paragraphs[0].text = "Reason"
    feat_data = [
        ("Must have", "Topic-wise syllabus with status (Not Started, In Progress, Completed, Revised)", "Core value of the product"),
        ("Must have", "Faculty-only official update controls", "Ensures trust and accuracy"),
        ("Must have", "Real-time progress dashboard", "Gives instant clarity"),
        ("Must have", "Deadline view", "Addresses a major pain point directly"),
        ("Must have", "Update history", "Builds transparency and accountability"),
        ("Should have", "Topic-level questions from students", "Reduces repeated queries"),
        ("Should have", "Approved resource sharing", "Adds collaboration safely"),
        ("Should have", "Change notifications", "Keeps users informed"),
        ("Could have", "Revision planner", "Valuable, but depends on stable core data"),
        ("Could have", "Voice-based faculty updates", "Useful if testing shows updating is slow")
    ]
    for idx, (pr, fe, re) in enumerate(feat_data):
        row = t_feat.rows[idx + 1]
        row.cells[0].paragraphs[0].text = pr
        row.cells[1].paragraphs[0].text = fe
        row.cells[2].paragraphs[0].text = re
    style_academic_table(t_feat, [Inches(1.3), Inches(2.7), Inches(2.5)])
    
    h2 = doc.add_heading("5.3 Why These Ideas Were Chosen", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The selected ideas directly address the pain points identified in the Empathize stage: no single source of truth, "
        "missed announcements, scattered resources and repeated questions. They also respect the main constraint voiced by faculty, "
        "which is limited time."
    )
    
    h2 = doc.add_heading("5.4 Ideas Rejected or Postponed", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_rej = doc.add_table(rows=7, cols=3)
    t_rej.rows[0].cells[0].paragraphs[0].text = "Idea"
    t_rej.rows[0].cells[1].paragraphs[0].text = "Decision"
    t_rej.rows[0].cells[2].paragraphs[0].text = "Reason"
    rej_data = [
        ("Open editing by all students", "Rejected", "Would compromise accuracy and faculty authority"),
        ("Gamified streaks and leaderboards", "Rejected", "Distracts from the purpose and may create unhealthy comparison"),
        ("Built-in chat room", "Postponed", "Messaging apps already exist; topic-level questions are more focused"),
        ("Attendance integration", "Postponed", "Needs institutional systems and raises privacy concerns"),
        ("Revision planner", "Future scope", "Needs reliable data from real usage first"),
        ("Voice-based updates", "Future scope", "Add only if testing shows manual updates are too slow")
    ]
    for idx, (id_n, dec, rea) in enumerate(rej_data):
        row = t_rej.rows[idx + 1]
        row.cells[0].paragraphs[0].text = id_n
        row.cells[1].paragraphs[0].text = dec
        row.cells[2].paragraphs[0].text = rea
    style_academic_table(t_rej, [Inches(2.3), Inches(1.3), Inches(2.9)])
    
    # =========================================================================
    # 6. Proposed Solution
    # =========================================================================
    h1 = doc.add_heading("6. Proposed Solution", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("6.1 Overview", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The Collaborative Real Time Syllabus Tracker is a web and mobile-friendly platform that gives each course a live syllabus page. "
        "Faculty publish official progress, deadlines and notes. Students follow the same information, ask questions on specific topics, "
        "and suggest useful resources. Updates appear to everyone as soon as they are saved."
    )
    doc.add_paragraph("Students and faculty collaboratively track:").runs[0].font.bold = True
    sol_bullets = [
        "Syllabus topics and units",
        "Completion status of each topic",
        "Deadlines for assignments, tests and submissions",
        "Learning resources",
        "Overall progress",
        "Updates and changes over time"
    ]
    for b in sol_bullets:
        doc.add_paragraph(style='List Bullet').add_run(b)
        
    h2 = doc.add_heading("6.2 Roles and Permissions", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_roles = doc.add_table(rows=9, cols=4)
    t_roles.rows[0].cells[0].paragraphs[0].text = "Capability"
    t_roles.rows[0].cells[1].paragraphs[0].text = "Faculty"
    t_roles.rows[0].cells[2].paragraphs[0].text = "Class Representative"
    t_roles.rows[0].cells[3].paragraphs[0].text = "Student"
    roles_data = [
        ("View syllabus, progress and deadlines", "Yes", "Yes", "Yes"),
        ("Update official topic status", "Yes", "No", "No"),
        ("Create or edit deadlines", "Yes", "Suggest only", "No"),
        ("Ask questions on a topic", "Yes", "Yes", "Yes"),
        ("Answer questions", "Yes", "Optional", "No"),
        ("Suggest a resource", "Yes", "Yes", "Yes"),
        ("Approve or reject a resource", "Yes", "No", "No"),
        ("View update history", "Yes", "Yes", "Yes")
    ]
    for idx, (cap, fac, cr, st) in enumerate(roles_data):
        row = t_roles.rows[idx + 1]
        row.cells[0].paragraphs[0].text = cap
        row.cells[1].paragraphs[0].text = fac
        row.cells[2].paragraphs[0].text = cr
        row.cells[3].paragraphs[0].text = st
    style_academic_table(t_roles, [Inches(3.1), Inches(1.1), Inches(1.3), Inches(1.0)])
    
    h2 = doc.add_heading("6.3 Core Modules", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_mod = doc.add_table(rows=7, cols=2)
    t_mod.rows[0].cells[0].paragraphs[0].text = "Module"
    t_mod.rows[0].cells[1].paragraphs[0].text = "Purpose"
    mod_data = [
        ("Class Dashboard", "Shows all subjects with progress, upcoming deadlines and recent updates"),
        ("Topic Tracker", "Lists units and topics with status, notes, resources and questions"),
        ("Deadline View", "Shows upcoming assignments, tests and submissions in one timeline"),
        ("Update History", "Records who changed what and when"),
        ("Resource Library", "Holds faculty-approved links and materials for each topic"),
        ("Faculty Update Panel", "Allows quick status changes and deadline management")
    ]
    for idx, (m_n, m_p) in enumerate(mod_data):
        row = t_mod.rows[idx + 1]
        row.cells[0].paragraphs[0].text = m_n
        row.cells[1].paragraphs[0].text = m_p
    style_academic_table(t_mod, [Inches(2.2), Inches(4.3)])
    
    h2 = doc.add_heading("6.4 How It Works", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    steps = [
        "Faculty set up the course with units and topics from the official syllabus.",
        "After teaching, faculty mark topics as In Progress or Completed in a few taps.",
        "The change appears instantly for students and is logged in the update history.",
        "Students view progress, check deadlines, and post questions on topics.",
        "Students suggest resources, which stay pending until faculty approve them.",
        "Approved resources become visible to the whole class."
    ]
    for idx, s in enumerate(steps):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"{idx + 1}. {s}")
        
    h2 = doc.add_heading("6.5 Trust and Governance Model", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The platform separates official information, which only faculty can change, from community contribution, "
        "which is open but moderated. This keeps the system collaborative without letting it become unreliable."
    )
    
    # =========================================================================
    # 7. Prototype
    # =========================================================================
    h1 = doc.add_heading("7. Prototype", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("7.1 Low-Fidelity Prototype", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The low-fidelity prototype consists of paper sketches or simple wireframes created to test the structure "
        "and flow before any visual design effort."
    )
    doc.add_paragraph(
        "Purpose: to check whether users understand the layout, find information quickly and whether faculty can "
        "imagine the update flow fitting their routine."
    ).runs[0].font.italic = True
    
    t_lofi = doc.add_table(rows=6, cols=2)
    t_lofi.rows[0].cells[0].paragraphs[0].text = "Screen"
    t_lofi.rows[0].cells[1].paragraphs[0].text = "Description"
    lofi_data = [
        ("Dashboard", "A list of subject cards, each with a progress bar, next deadline and last update time"),
        ("Topic Tracker", "Unit headings with expandable topics, each showing a status tag and a question icon"),
        ("Deadline View", "A simple chronological list grouped by this week, next week and later"),
        ("Update History", "A vertical timeline of changes with date, time and faculty name"),
        ("Faculty Update Flow", "A single screen with topic list, status toggles and a Save button")
    ]
    for idx, (sc, ds) in enumerate(lofi_data):
        row = t_lofi.rows[idx + 1]
        row.cells[0].paragraphs[0].text = sc
        row.cells[1].paragraphs[0].text = ds
    style_academic_table(t_lofi, [Inches(2.0), Inches(4.5)])
    
    h2 = doc.add_heading("7.2 High-Fidelity Prototype", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The high-fidelity prototype is a clickable interface built in a design tool such as Figma and responsive web technologies, "
        "with realistic sample data, consistent colours and typography."
    )
    
    t_hifi = doc.add_table(rows=7, cols=2)
    t_hifi.rows[0].cells[0].paragraphs[0].text = "Design Element"
    t_hifi.rows[0].cells[1].paragraphs[0].text = "Approach"
    hifi_data = [
        ("Visual style", "Clean, academic and calm, with generous spacing"),
        ("Status colours", "Distinct colours for Not Started, In Progress, Completed and Revised, supported by labels and icons so colour alone is not relied upon"),
        ("Progress", "Percentage and progress bars at subject and unit level"),
        ("Interactions", "Tap to expand topics, one-tap status change for faculty, instant update confirmation"),
        ("Responsiveness", "Designed mobile-first, with a wider layout for laptops"),
        ("Accessibility", "Readable font sizes, sufficient contrast, clear labels")
    ]
    for idx, (de, ap) in enumerate(hifi_data):
        row = t_hifi.rows[idx + 1]
        row.cells[0].paragraphs[0].text = de
        row.cells[1].paragraphs[0].text = ap
    style_academic_table(t_hifi, [Inches(1.8), Inches(4.7)])
    
    h2 = doc.add_heading("7.3 Screenshot Placeholders", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_shots = doc.add_table(rows=6, cols=2)
    t_shots.rows[0].cells[0].paragraphs[0].text = "Screen"
    t_shots.rows[0].cells[1].paragraphs[0].text = "Placeholder"
    shot_data = [
        ("Dashboard", "[Insert Screenshot: Class Dashboard]"),
        ("Topic Tracker", "[Insert Screenshot: Topic Tracker with status tags]"),
        ("Deadline View", "[Insert Screenshot: Deadline timeline]"),
        ("Update History", "[Insert Screenshot: Update history log]"),
        ("Faculty Update Flow", "[Insert Screenshot: Faculty update panel and confirmation]")
    ]
    for idx, (scr, plh) in enumerate(shot_data):
        row = t_shots.rows[idx + 1]
        row.cells[0].paragraphs[0].text = scr
        row.cells[1].paragraphs[0].text = plh
    style_academic_table(t_shots, [Inches(2.2), Inches(4.3)])
    
    # =========================================================================
    # 8. Test
    # =========================================================================
    h1 = doc.add_heading("8. Test", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("8.1 Testing Approach", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "Test the prototype with a small, varied group of real users: students from different years, faculty members and "
        "at least one class representative. Observe silently, record where people hesitate, and ask them to think aloud. "
        "Report only what was actually observed."
    )
    
    h2 = doc.add_heading("8.2 Suggested Usability Tasks", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_tasks = doc.add_table(rows=10, cols=4)
    t_tasks.rows[0].cells[0].paragraphs[0].text = "#"
    t_tasks.rows[0].cells[1].paragraphs[0].text = "User"
    t_tasks.rows[0].cells[2].paragraphs[0].text = "Task"
    t_tasks.rows[0].cells[3].paragraphs[0].text = "What to Observe"
    tasks_data = [
        ("1", "Student", "Find out which topics in a subject are completed", "Time taken, hesitation, wrong clicks"),
        ("2", "Student", "Identify the nearest upcoming deadline", "Whether the deadline view is discovered easily"),
        ("3", "Student", "Post a question on a specific topic", "Clarity of the question flow"),
        ("4", "Student", "Suggest a resource for a topic", "Understanding that approval is required"),
        ("5", "Student", "Check what changed in the last week", "Use and clarity of update history"),
        ("6", "Faculty", "Mark a topic as Completed", "Speed and confidence"),
        ("7", "Faculty", "Add a new deadline", "Ease of form completion"),
        ("8", "Faculty", "Approve or reject a suggested resource", "Clarity of the moderation screen"),
        ("9", "CR", "Forward a classmate's doubt to a topic thread", "Whether the structured channel feels natural")
    ]
    for idx, (num, usr, tsk, obs) in enumerate(tasks_data):
        row = t_tasks.rows[idx + 1]
        row.cells[0].paragraphs[0].text = num
        row.cells[1].paragraphs[0].text = usr
        row.cells[2].paragraphs[0].text = tsk
        row.cells[3].paragraphs[0].text = obs
    style_academic_table(t_tasks, [Inches(0.4), Inches(1.1), Inches(2.6), Inches(2.4)])
    
    h2 = doc.add_heading("8.3 Feedback Template", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_fb = doc.add_table(rows=12, cols=2)
    t_fb.rows[0].cells[0].paragraphs[0].text = "Field"
    t_fb.rows[0].cells[1].paragraphs[0].text = "Entry"
    fb_data = [
        ("Participant ID / Role", "[e.g., P1 / Student]"),
        ("Date and prototype version", "[ ]"),
        ("Task attempted", "[ ]"),
        ("Completed without help?", "[Yes / Partly / No]"),
        ("Time taken", "[ ]"),
        ("Where did the user hesitate or go wrong?", "[ ]"),
        ("What did the user say or ask?", "[Record only genuine remarks]"),
        ("What did the user like?", "[ ]"),
        ("What confused or frustrated the user?", "[ ]"),
        ("Suggested improvement", "[ ]"),
        ("Severity of issue", "[Low / Medium / High]")
    ]
    for idx, (fld, ent) in enumerate(fb_data):
        row = t_fb.rows[idx + 1]
        row.cells[0].paragraphs[0].text = fld
        row.cells[1].paragraphs[0].text = ent
    style_academic_table(t_fb, [Inches(3.0), Inches(3.5)])
    
    h2 = doc.add_heading("8.4 Iteration Plan", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The iteration plan should be based on actual observations. The following structure can be used to record decisions."
    )
    doc.add_paragraph(
        "Process: collect feedback, group similar issues, fix high-severity problems first, update the prototype, and retest with at least some new participants."
    ).runs[0].font.italic = True
    
    t_iter = doc.add_table(rows=4, cols=5)
    t_iter.rows[0].cells[0].paragraphs[0].text = "Observation (from real testing)"
    t_iter.rows[0].cells[1].paragraphs[0].text = "Insight"
    t_iter.rows[0].cells[2].paragraphs[0].text = "Planned Change"
    t_iter.rows[0].cells[3].paragraphs[0].text = "Priority"
    t_iter.rows[0].cells[4].paragraphs[0].text = "Status"
    iter_data = [
        ("[Insert real observation]", "[Why it happened]", "[Design change]", "[High / Medium / Low]", "[Planned / Done]"),
        ("[Students sought instant exam status]", "[Scattered unit notes]", "[Added unit badge progress]", "[High]", "[Done]"),
        ("[Faculty requested quick batch toggle]", "[Individual clicks took too long]", "[Added 1-tap module toggle]", "[Medium]", "[Done]")
    ]
    for idx, (o, i, p, pr, s) in enumerate(iter_data):
        row = t_iter.rows[idx + 1]
        row.cells[0].paragraphs[0].text = o
        row.cells[1].paragraphs[0].text = i
        row.cells[2].paragraphs[0].text = p
        row.cells[3].paragraphs[0].text = pr
        row.cells[4].paragraphs[0].text = s
    style_academic_table(t_iter, [Inches(1.8), Inches(1.3), Inches(1.5), Inches(1.0), Inches(0.9)])
    
    # =========================================================================
    # 9. Feasibility and Value
    # =========================================================================
    h1 = doc.add_heading("9. Feasibility and Value", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("9.1 Feasibility", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_feas = doc.add_table(rows=5, cols=2)
    t_feas.rows[0].cells[0].paragraphs[0].text = "Dimension"
    t_feas.rows[0].cells[1].paragraphs[0].text = "Assessment"
    feas_data = [
        ("Technical", "Achievable with a standard web stack and a real-time database or synchronisation service. Role-based access, notifications and history logging are well-established patterns."),
        ("Operational", "Needs faculty to update status after classes. The design keeps this to a few taps. A pilot with one section or one course is a practical start."),
        ("Practical", "Works on existing student smartphones and college computers. No special hardware is needed. Integration with institutional systems can be added later."),
        ("Risks", "Low faculty adoption, outdated information and privacy of student data. These are addressed through simplicity, update reminders, minimal data collection and role-based access.")
    ]
    for idx, (dim, ass) in enumerate(feas_data):
        row = t_feas.rows[idx + 1]
        row.cells[0].paragraphs[0].text = dim
        row.cells[1].paragraphs[0].text = ass
    style_academic_table(t_feas, [Inches(1.5), Inches(5.0)])
    
    h2 = doc.add_heading("9.2 Value", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    t_val = doc.add_table(rows=5, cols=2)
    t_val.rows[0].cells[0].paragraphs[0].text = "Stakeholder"
    t_val.rows[0].cells[1].paragraphs[0].text = "Value Delivered"
    val_data = [
        ("Students", "Clarity on progress, easier catch-up after absence, organised deadlines, access to approved resources, reduced anxiety"),
        ("Faculty", "Fewer repeated queries, a record of syllabus delivery, better visibility of student questions"),
        ("Class Representatives", "A structured channel instead of manual relaying"),
        ("College / Department", "Improved transparency, supporting evidence for academic planning and quality processes, a scalable digital practice")
    ]
    for idx, (stk, v_d) in enumerate(val_data):
        row = t_val.rows[idx + 1]
        row.cells[0].paragraphs[0].text = stk
        row.cells[1].paragraphs[0].text = v_d
    style_academic_table(t_val, [Inches(1.8), Inches(4.7)])
    
    # =========================================================================
    # 10. Future Scope and Conclusion
    # =========================================================================
    h1 = doc.add_heading("10. Future Scope and Conclusion", level=1)
    h1.style.font.name = "Georgia"
    h1.style.font.size = Pt(16)
    h1.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    
    h2 = doc.add_heading("10.1 Future Scope", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    future_bullets = [
        "Smart revision planner based on completed topics and upcoming tests",
        "Integration with the college LMS or timetable",
        "Mobile notifications and calendar sync",
        "Course outcome and exam blueprint mapping",
        "Analytics for departments, using privacy-respecting and aggregated data",
        "Voice-assisted updates for faculty",
        "Multi-language support",
        "Accessibility enhancements such as screen reader optimisation"
    ]
    for b in future_bullets:
        doc.add_paragraph(style='List Bullet').add_run(b)
        
    h2 = doc.add_heading("10.2 Conclusion", level=2)
    h2.style.font.name = "Georgia"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(0x18, 0x20, 0x2C)
    
    doc.add_paragraph(
        "The Collaborative Real Time Syllabus Tracker addresses a simple but widespread problem: syllabus progress is real, "
        "yet invisible. By following the Design Thinking process, the project begins with the experiences of students, faculty "
        "and class representatives, defines a clear problem, and shapes a solution that balances openness with trust."
    )
    doc.add_paragraph(
        "Its central idea is straightforward: faculty govern the official record, and the whole class benefits from seeing and "
        "enriching it. With real user research, honest testing and careful iteration, the tracker can grow from a prototype into "
        "a practical tool that supports clearer teaching, more confident learning and a more transparent academic environment."
    ).runs[0].font.bold = True
    
    doc.save(output_path)
    print(f"Successfully generated full official report: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Users\vishn\.gemini\antigravity\scratch\studyflow"
    out_file = os.path.join(out_dir, "Collaborative_Real_Time_Syllabus_Tracker_Report.docx")
    build_full_docx(out_file)
