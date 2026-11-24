"""Builds the project report (PhotoShield_Report.docx). Run: python docs/report/build_report.py"""
import os

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.dirname(HERE)
OUT = os.path.join(HERE, "PhotoShield_Report.docx")
ACCENT = RGBColor(0x2B, 0x3A, 0x8C)
FONT = "Calibri"


def shade(cell, hex_fill):
    props = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    props.append(shd)


def set_cell_margins(table, top=50, bottom=50, left=90, right=90):
    tbl_pr = table._tbl.tblPr
    margins = OxmlElement("w:tblCellMar")
    for side, value in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        margins.append(el)
    tbl_pr.append(margins)


def add_page_number(paragraph):
    run = paragraph.add_run()
    for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def para(doc, text="", bold_lead=None, style=None, after=4, align=None, size=None, italic=False):
    p = doc.add_paragraph(style=style)
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
    r = p.add_run(text)
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(after)
    if align:
        p.alignment = align
    return p


def bullets(doc, items):
    for item in items:
        if isinstance(item, tuple):
            para(doc, item[1], bold_lead=item[0], style="List Bullet", after=1)
        else:
            para(doc, item, style="List Bullet", after=1)
    doc.paragraphs[-1].paragraph_format.space_after = Pt(5)


def table(doc, header, rows, widths_cm, font_size=9.5):
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_margins(t)
    for row_idx, values in enumerate([header] + rows):
        tr_pr = t.rows[row_idx]._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        for col_idx, value in enumerate(values):
            cell = t.rows[row_idx].cells[col_idx]
            cell.width = Cm(widths_cm[col_idx])
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(value)
            run.font.size = Pt(font_size)
            # Keep the table on one page: rows can't split, and every row but the last stays with the next.
            p.paragraph_format.keep_with_next = row_idx < len(rows)
            if row_idx == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                shade(cell, "2B3A8C")
            elif row_idx % 2 == 0:
                shade(cell, "EEF1FA")
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(doc, path, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Cm(width_cm))
    c = para(doc, caption, align=WD_ALIGN_PARAGRAPH.CENTER, size=9, italic=True, after=8)
    return c


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.2)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.paragraph_format.line_spacing = 1.08
    for name, size, before in (("Heading 1", 14, 10), ("Heading 2", 12, 6)):
        st = doc.styles[name]
        st.font.name = FONT
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = ACCENT
        st.element.rPr.rFonts.set(qn("w:asciiTheme"), "") if False else None
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(4)
        st.paragraph_format.keep_with_next = True

    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run("UCS531 Cloud Computing · PhotoShield · Page ")
    r.font.size = Pt(9)
    add_page_number(footer)

    # ---- Title block ----
    t = para(doc, "PhotoShield: A High-Availability Photo-Privacy Service on AWS",
             align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    t.runs[0].bold = True
    t.runs[0].font.size = Pt(18)
    t.runs[0].font.color.rgb = ACCENT
    para(doc, "Project report · UCS531 Cloud Computing · Thapar Institute of Engineering & Technology, Patiala",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=10, after=1)
    para(doc, "Team: Gavisht · Aarpit · Nishant          November 2026",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=10, after=10)

    # ---- 1 ----
    doc.add_heading("1. Problem statement and objectives", level=1)
    para(doc, "A web application on a single server is a single point of failure: a crashed instance, a failed "
              "data centre or a traffic spike takes it offline. Our proposal was to build a web application on AWS "
              "that stays available through failures and scales with load, using only free-tier-eligible services.")
    para(doc, "As the workload we built PhotoShield, a real privacy tool. Photos from phones, cameras and drones carry "
              "hidden EXIF/XMP metadata: exact GPS coordinates, the device and its serial number, and the time of "
              "capture. Anyone who receives the original file can read it. PhotoShield shows the user what a photo "
              "reveals and returns a clean copy. Image processing is CPU-intensive, so heavy use genuinely drives "
              "the auto-scaling mechanism rather than requiring an artificial load.")
    para(doc, "Objectives:", after=1).runs[0].bold = True
    bullets(doc, [
        "No single point of failure in the web tier: at least two instances in two Availability Zones (AZs).",
        "Self-healing: failed instances are detected and replaced automatically.",
        "Automatic horizontal scaling driven by CPU utilisation.",
        "Secure by default: private database, no credentials in code or on servers, least-privilege access.",
        "Fully reproducible infrastructure (Infrastructure as Code) that can be destroyed when idle to save cost.",
        "Measure availability, scaling and recovery with real tests.",
    ])

    # ---- 2 ----
    doc.add_heading("2. System architecture", level=1)
    figure(doc, os.path.join(DOCS, "architecture.png"), 16.5,
           "Figure 1. Architecture. Arrows show request flow; labels give the port allowed by each security group.")
    para(doc, "A user's request reaches the internet-facing Application Load Balancer (ALB), which forwards it to a "
              "healthy instance in either AZ. The instance processes the photo, stores the clean copy in S3, records "
              "the scan in RDS MySQL and returns the result page. The table summarises each component.")
    table(doc, ["Layer", "Service", "Configuration"], [
        ["Load balancing", "Application Load Balancer", "HTTP :80; health check GET /health every 15 s (2 failures = unhealthy)"],
        ["Compute", "EC2 Auto Scaling group", "t3.micro, Amazon Linux 2023; min 2 / max 4 across us-east-1a and 1b; target tracking at 50% average CPU; ELB health checks; rolling instance refresh"],
        ["Database", "RDS MySQL 8.4", "db.t3.micro, not publicly accessible, encrypted storage, daily backups, TLS connections"],
        ["Storage", "Amazon S3", "Private bucket, encryption at rest, HTTPS-only policy, pre-signed downloads, 24-hour expiry of cleaned photos"],
        ["Security", "IAM, Secrets Manager, SSM", "Instance role (least privilege), DB password generated and rotated by RDS, Session Manager instead of SSH, IMDSv2"],
        ["Monitoring", "CloudWatch, SNS", "1-minute metrics, dashboard, alarms (unhealthy hosts, 5xx errors, DB CPU, DB storage)"],
    ], [2.6, 3.8, 10.2])

    # ---- 3 ----
    doc.add_heading("3. Implementation", level=1)
    doc.add_heading("3.1 Infrastructure as Code", level=2)
    para(doc, "All 39 AWS resources are defined in Terraform (AWS provider 6.x), split by concern: network, IAM, "
              "storage, database, load balancer, compute and monitoring. Terraform derives the creation order from "
              "references between resources. Every resource is tagged Project=cloud-ha and ManagedBy=terraform. "
              "Two scripts wrap the workflow: up.ps1 builds the stack and waits until every target is healthy "
              "(11-14 minutes, dominated by RDS creation), and down.ps1 destroys it. Because a rebuild takes one "
              "command, the stack is destroyed whenever it is not in use.")
    doc.add_heading("3.2 Application", level=2)
    para(doc, "PhotoShield is a Python Flask application served by gunicorn on port 8000 under an unprivileged "
              "system user. The image module (Pillow) reads GPS coordinates from EXIF and XMP, converting "
              "degrees/minutes/seconds to decimal, plus device, lens, serial number, time and software, and rates the "
              "privacy risk as high, medium or low. Cleaning applies the EXIF rotation to the pixels and then builds "
              "a new image from raw pixel data only, so no metadata can be carried over. Uploads are limited to "
              "15 MB and images above 40 megapixels are refused before decoding (decompression-bomb protection). "
              "The clean copy goes to S3 and is offered through a 10-minute pre-signed URL. Each scan is recorded in "
              "RDS with counts and yes/no flags only: coordinates are never stored. SQL uses parameterised queries. "
              "27 automated tests cover the image logic and the web routes with S3 and MySQL replaced by fakes.")
    figure(doc, os.path.join(DOCS, "results", "screenshot-result.png"), 13.0,
           "Figure 2. Result page for the sample photo: location, device and time found, with the clean copy from S3. "
           "The header shows which instance and AZ served the request.")
    doc.add_heading("3.3 Security controls", level=2)
    bullets(doc, [
        ("Network isolation: ", "security groups reference each other, so the ALB accepts port 80 from the "
                                "internet, instances accept 8000 only from the ALB, and RDS accepts 3306 only from instances."),
        ("No stored credentials: ", "RDS generates the database password and keeps it in Secrets Manager; "
                                    "instances read it at runtime through their IAM role and re-fetch it after rotation."),
        ("Least privilege: ", "instances may read the application code but write only to the clean/ prefix, so a "
                              "compromised server cannot replace the code."),
        ("Hardening: ", "IMDSv2 only (blocks SSRF theft of role credentials), no SSH or key pairs (Session Manager), "
                        "TLS to RDS with certificate verification, encryption at rest for RDS and S3."),
    ])
    doc.add_heading("3.4 Deployment of new versions", level=2)
    para(doc, "Terraform uploads the application to S3 and every new instance downloads it at boot (EC2 user data "
              "is limited to 16 KB). A hash of the application files is part of the launch template, so a code "
              "change creates a new template version and triggers a rolling instance refresh that replaces one "
              "instance at a time while at least 50% keep serving.")

    # ---- 4 ----
    doc.add_heading("4. Testing and results", level=1)
    para(doc, "All tests ran against the live deployment in us-east-1 from a client in India. The load and failover "
              "tests are scripts in the repository (loadtest.py, failover_test.py); graphs come from CloudWatch."
         ).paragraph_format.keep_with_next = True
    table(doc, ["Test", "Method", "Result"], [
        ["Fresh deployment", "up.ps1 until all targets healthy", "11-14 min"],
        ["Rolling update", "New app version, both instances replaced; site probed every 10 s", "14.5 min; 79/79 requests OK, 0 failed"],
        ["Load test", "24 concurrent clients, 15 min of photo scans (6 MP each)", "6,676 requests, 0 failed; 2 → 4 instances; 4.7 → 9.6 req/s; median latency 5.0 s → 1.7-2.5 s"],
        ["Failover", "Terminate 1 of 2 instances; site probed every second", "Replacement healthy after 184 s (3.1 min); 5 of 177 requests failed, all within 21 s"],
        ["Processing time", "Single request on t3.micro", "~465 ms for a 6 MP photo; 45-90 ms for 1-1.6 MP"],
    ], [3.0, 6.6, 7.0])
    figure(doc, os.path.join(DOCS, "results", "scaling.png"), 14.5,
           "Figure 3. Load test (CloudWatch, 1-minute data): CPU, instances in service and requests per minute.")
    para(doc, "Scaling. ", bold_lead=None).runs[0].bold = True
    doc.paragraphs[-1].add_run(
        "Average CPU rose to about 90% within two minutes. The target-tracking policy raised the desired capacity "
        "from 2 to 4 about 4.6 minutes after the test started, and the new instances began serving roughly two "
        "minutes later. Throughput doubled while no request failed. CPU then stayed near 88% because the group had "
        "reached its configured maximum of four instances: the maximum works as a cost ceiling, and raising it would "
        "add capacity. Enabling 1-minute (detailed) monitoring was necessary; with default 5-minute metrics the "
        "policy would react only after about 15 minutes.")
    p = para(doc, "", after=4)
    p.add_run("Failover. ").bold = True
    p.add_run("Terminating an instance abruptly caused five failed requests in the first 21 seconds: until the ALB's "
              "health check marks the instance unhealthy (two failed checks, 15 s apart), some requests are still "
              "routed to it. Planned replacements behave differently: during the rolling update instances were "
              "deregistered before termination and no request failed. The Auto Scaling group launched a replacement "
              "after 85 s and restored full capacity after 3.1 minutes. A shorter health-check interval would narrow "
              "the window of errors at the cost of more health-check traffic.")

    # ---- 5 ----
    doc.add_heading("5. Cost", level=1)
    para(doc, "The stack costs about USD 2 per day while running, mostly the load balancer and the database. It is "
              "destroyed after each session, so development and all tests consumed only a few dollars of AWS "
              "credits. Cost-driven choices: t3.micro instances, single-AZ RDS, the default VPC's public subnets "
              "instead of private subnets with a NAT gateway (about USD 32 per month), a 24-hour expiry on stored "
              "photos, and a maximum of four instances.")

    # ---- 6 ----
    doc.add_heading("6. Limitations and future work", level=1)
    bullets(doc, [
        ("Database availability: ", "RDS is single-AZ and is the remaining single point of failure. Setting "
                                    "multi_az = true adds a synchronous standby with automatic failover (about double the DB cost)."),
        ("HTTPS: ", "the site uses HTTP; a domain with an ACM certificate on the ALB, or CloudFront, would add TLS for users."),
        ("Network: ", "private subnets with VPC endpoints would remove public IPs from the instances."),
        ("Abuse protection: ", "AWS WAF rate limiting; an SQS queue with workers for very large images."),
        ("Automation: ", "a CI/CD pipeline that runs the tests and applies Terraform on every commit."),
    ])

    # ---- 7 ----
    doc.add_heading("7. Conclusion", level=1)
    para(doc, "PhotoShield meets the proposal's goals on free-tier-sized resources: the web tier has no single point "
              "of failure across two AZs, replaces failed instances in about three minutes, doubles its capacity "
              "automatically under load, and rolls out new versions with zero failed requests. The whole system is "
              "reproducible from code and secure by default, and the application addresses a real privacy problem. "
              "The main remaining weakness, the single-AZ database, is a deliberate cost trade-off with a one-line fix.")

    doc.add_heading("References", level=1)
    for ref in [
        "AWS. Elastic Load Balancing: health checks for Application Load Balancer target groups. docs.aws.amazon.com",
        "AWS. Amazon EC2 Auto Scaling: target tracking scaling policies; instance refresh. docs.aws.amazon.com",
        "AWS. Amazon RDS: Multi-AZ deployments; password management with AWS Secrets Manager. docs.aws.amazon.com",
        "AWS. Amazon EC2: Instance Metadata Service Version 2; burstable performance instances. docs.aws.amazon.com",
        "HashiCorp. Terraform AWS Provider documentation. registry.terraform.io/providers/hashicorp/aws",
        "CIPA DC-008. Exchangeable image file format for digital still cameras: Exif Version 2.32.",
    ]:
        para(doc, ref, style="List Number", size=9.5, after=0)

    doc.save(OUT)
    print("Saved", OUT)


if __name__ == "__main__":
    build()
