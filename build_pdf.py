import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

def generate_pdf():
    pdf_filename = r"d:\MOBILE RECHARGE APPLICATION\funngro\Portfolio_Website_Submission_Bapanapalli_Shankar.pdf"
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()

    # Custom color scheme
    primary_color = colors.HexColor("#059669")   # Emerald green
    dark_color = colors.HexColor("#0f172a")      # Dark slate
    secondary_color = colors.HexColor("#2563eb") # Royal blue
    light_bg = colors.HexColor("#f8fafc")

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=dark_color,
        alignment=1
    )

    subtitle_style = ParagraphStyle(
        'DocSubtitle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=primary_color,
        alignment=1,
        spaceAfter=10
    )

    h2_style = ParagraphStyle(
        'SectionH2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=17,
        textColor=primary_color,
        spaceBefore=14,
        spaceAfter=6
    )

    body_style = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.HexColor("#334155")
    )

    bullet_style = ParagraphStyle(
        'BulletCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#334155"),
        leftIndent=12
    )

    link_style = ParagraphStyle(
        'LinkCustom',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13.5,
        textColor=secondary_color
    )

    story = []

    # Title & Subtitle Header
    story.append(Paragraph("FUNNGRO PROJECT SUBMISSION REPORT", title_style))
    story.append(Paragraph("Portfolio Website Creation Project & Technical Documentation", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=2, color=primary_color, spaceAfter=8))

    # Candidate Header Table
    candidate_data = [
        [Paragraph("<b>Candidate Name:</b> Bapanapalli Shankar", body_style), Paragraph("<b>Target Role:</b> Full Stack & AI Data Analyst", body_style)],
        [Paragraph("<b>Email:</b> shankarcode291439@gmail.com", body_style), Paragraph("<b>Phone:</b> +91 7842161185", body_style)],
        [Paragraph("<b>GitHub:</b> github.com/Shankarz9704", body_style), Paragraph("<b>LinkedIn:</b> linkedin.com/in/shankar9704", body_style)]
    ]
    t_candidate = Table(candidate_data, colWidths=[270, 270])
    t_candidate.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), light_bg),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_candidate)
    story.append(Spacer(1, 8))

    # Website Links Section
    story.append(Paragraph("🌐 Website Access & Live Links", h2_style))
    links_data = [
        [Paragraph("<b>Live Public Website URL:</b>", body_style), Paragraph("<font color='#2563eb'><u>https://1e87a5dc583bf0.lhr.life</u></font>", link_style)],
        [Paragraph("<b>Local Web Server URL:</b>", body_style), Paragraph("<font color='#2563eb'><u>http://localhost:8080</u></font>", link_style)],
        [Paragraph("<b>GitHub Repository:</b>", body_style), Paragraph("<font color='#2563eb'><u>https://github.com/Shankarz9704/PORTFOLIO</u></font>", link_style)],
        [Paragraph("<b>Project Folder Path:</b>", body_style), Paragraph("d:\\MOBILE RECHARGE APPLICATION\\funngro", body_style)]
    ]
    t_links = Table(links_data, colWidths=[160, 380])
    t_links.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#ecfdf5")),
        ('BOX', (0,0), (-1,-1), 1, primary_color),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#a7f3d0")),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_links)
    story.append(Spacer(1, 10))

    # Uploaded Screenshots Section
    story.append(Paragraph("📸 Complete Live Website Screenshots", h2_style))

    uploaded_base = r"C:\Users\B SHANKAR\.gemini\antigravity\brain\cc89494d-8e14-4527-86a6-1f1033da9ac2\.user_uploaded"
    
    img_hero = os.path.join(uploaded_base, "media_1787739360261.png")
    img_about = os.path.join(uploaded_base, "media_1787739382901.png")
    img_exp = os.path.join(uploaded_base, "media_1787739336529.png")
    img_skills = os.path.join(uploaded_base, "media_1787739407894.png")
    img_projects1 = os.path.join(uploaded_base, "media_1787739440012.png")
    img_projects2 = os.path.join(uploaded_base, "media_1787739481900.png")
    img_edu = os.path.join(uploaded_base, "media_1787739497943.png")
    img_contact = os.path.join(uploaded_base, "media_1787739512380.png")

    # Figure 1: Hero Section
    if os.path.exists(img_hero):
        story.append(Paragraph("<b>Figure 1:</b> Hero Section UI (Shankar.dev Header & Branding)", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_hero, width=7.2*inch, height=2.4*inch))
        story.append(Spacer(1, 10))

    # Figure 2: About Me
    if os.path.exists(img_about):
        story.append(Paragraph("<b>Figure 2:</b> About Me & Key Metrics (5+ Projects, 6+ Certifications, 75% B.Tech)", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_about, width=7.2*inch, height=3.6*inch))
        story.append(Spacer(1, 10))

    # Figure 3: Experience Timeline
    if os.path.exists(img_exp):
        story.append(Paragraph("<b>Figure 3:</b> Professional & Internship Experience Timeline", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_exp, width=7.2*inch, height=3.8*inch))
        story.append(Spacer(1, 10))

    # Figure 4: Skills & Competencies
    if os.path.exists(img_skills):
        story.append(Paragraph("<b>Figure 4:</b> Technical Toolkit & Skill Bars (Development, Analytics, Databases)", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_skills, width=7.2*inch, height=3.0*inch))
        story.append(Spacer(1, 10))

    # Figure 5: Featured Projects Grid 1
    if os.path.exists(img_projects1):
        story.append(Paragraph("<b>Figure 5:</b> Featured Projects Gallery Grid 1 (AI Analytics, Security, Mobile Recharge)", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_projects1, width=7.2*inch, height=3.4*inch))
        story.append(Spacer(1, 10))

    # Figure 6: Featured Projects Grid 2
    if os.path.exists(img_projects2):
        story.append(Paragraph("<b>Figure 6:</b> Featured Projects Gallery Grid 2 (Food Order System & Student Registration)", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_projects2, width=7.2*inch, height=3.4*inch))
        story.append(Spacer(1, 10))

    # Figure 7: Education & Certifications
    if os.path.exists(img_edu):
        story.append(Paragraph("<b>Figure 7:</b> Education History & Industry Certifications (IBM, Cisco, Coursera, MongoDB)", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_edu, width=7.2*inch, height=3.6*inch))
        story.append(Spacer(1, 10))

    # Figure 8: Contact Section
    if os.path.exists(img_contact):
        story.append(Paragraph("<b>Figure 8:</b> Contact Information Card & Interactive Message Form", body_style))
        story.append(Spacer(1, 3))
        story.append(Image(img_contact, width=7.2*inch, height=3.6*inch))
        story.append(Spacer(1, 10))

    # Overview & Requirements
    story.append(Paragraph("📋 Project Description & Criteria Fulfilled", h2_style))
    story.append(Paragraph("This project delivers a responsive, modern personal portfolio website built to showcase Bapanapalli Shankar's technical skills, full stack projects, AI data analytics experience, and industry credentials.", body_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("• <b>Clean Design & Intuitive Navigation:</b> Sticky navigation bar with smooth scrolling to Home, About, Experience, Skills, Projects, Education, and Contact.", bullet_style))
    story.append(Paragraph("• <b>Mobile-Friendly Layout:</b> Responsive CSS flexbox/grid layout rendering across mobile, tablet, and desktop viewports.", bullet_style))
    story.append(Paragraph("• <b>Dark / Light Theme Engine:</b> One-click theme switcher with persistent local storage.", bullet_style))
    story.append(Paragraph("• <b>Interactive Form:</b> Real-time client contact form with input validation and response toast.", bullet_style))
    story.append(Spacer(1, 10))

    # Projects Table
    story.append(Paragraph("🚀 Featured Projects Showcase", h2_style))
    proj_table_data = [
        [Paragraph("<b>Project Name</b>", body_style), Paragraph("<b>Tech Stack</b>", body_style), Paragraph("<b>Description</b>", body_style)],
        [
            Paragraph("<b>Mobile Recharge Application</b>", body_style),
            Paragraph("Java 17, Spring Boot, JPA, REST API", body_style),
            Paragraph("Full-stack recharge management platform with payment flow and REST endpoints.", body_style)
        ],
        [
            Paragraph("<b>AI Data Analytics Dashboard</b>", body_style),
            Paragraph("Power BI, Python, Pandas, NumPy, SQL", body_style),
            Paragraph("Business intelligence dashboard for operational metrics and data visualization.", body_style)
        ],
        [
            Paragraph("<b>Online Food Order System</b>", body_style),
            Paragraph("Java 17, Spring Boot, DTO Pattern, JUnit", body_style),
            Paragraph("Enterprise food ordering backend API supporting status tracking & patch updates.", body_style)
        ],
        [
            Paragraph("<b>Student Course Registration</b>", body_style),
            Paragraph("Java, Spring Boot, JPA, HTML5", body_style),
            Paragraph("University course enrollment management portal and student database.", body_style)
        ],
        [
            Paragraph("<b>Secure Video System (AES)</b>", body_style),
            Paragraph("Python, AES Cryptography, OOP", body_style),
            Paragraph("Cryptographic system for secure video encryption, storage, and authentication.", body_style)
        ]
    ]
    t_proj = Table(proj_table_data, colWidths=[150, 140, 250])
    t_proj.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_color),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_proj)
    story.append(Spacer(1, 10))

    # Experience & Education Summary
    story.append(Paragraph("🎓 Education & Experience Summary", h2_style))
    story.append(Paragraph("• <b>Education:</b> B.Tech in Computer Science & Engineering (2021 – 2025, 75.0%) - Sri Venkateswara Engineering College, Tirupati.", bullet_style))
    story.append(Paragraph("• <b>Experience:</b> Concentrix (Advisor I), Altruist Technologies (Business Associate), Rooman Technologies (AI Data Analyst Intern), TaPTaP Blackbucks (AI/ML Intern), ExcelR (Python Full Stack Intern).", bullet_style))
    story.append(Paragraph("• <b>Certifications:</b> IBM (Python 101, SQL Databases), Cisco (PCAP Python, CCNA Networks), Coursera (Excel), MongoDB.", bullet_style))
    story.append(Spacer(1, 12))

    # Sign-off Footer Box
    sign_data = [
        [Paragraph("<b>Funngro Project Status:</b> Completed & Verified", body_style), Paragraph("<b>Submission Date:</b> August 26, 2026", body_style)]
    ]
    t_sign = Table(sign_data, colWidths=[270, 270])
    t_sign.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_sign)

    doc.build(story)
    print(f"PDF generated successfully at: {pdf_filename}")

if __name__ == '__main__':
    generate_pdf()
