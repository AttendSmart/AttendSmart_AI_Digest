import customtkinter as ctk
from tkinter import messagebox
from tkinter.filedialog import asksaveasfilename
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from google import genai
import difflib

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch


# ---------------- GOOGLE SHEETS LOGIN ----------------
SCOPE = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
CREDS_FILE = "attendsmart.json"
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
gspread_client = gspread.authorize(creds)

MARKS_BOOK = "Student Performance 2025"
MARKS_SHEET = "performance"

marks_ws = gspread_client.open(MARKS_BOOK).worksheet(MARKS_SHEET)
marks_df = pd.DataFrame(marks_ws.get_all_records())


# ---------------- SUMMARY FUNCTIONS ------------------
def summarize_student(student_name: str, df: pd.DataFrame):
    """Summarize student with structured data."""
    rec = df[df['Name'].str.lower() == student_name.lower()]
    if rec.empty:
        return None, f"No record found for '{student_name}'."

    row = rec.iloc[0]
    numeric_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
    text_cols = [col for col in df.columns if col not in numeric_cols + ['Name']]

    # Collect structured info
    student_info = {
        "Name": row.get("Name", "-"),
        "Grade": row.get("Grade", "-"),
    }

    # Marks
    marks = {}
    for col in numeric_cols:
        marks[col] = row[col]
    if marks:
        avg = sum(marks.values()) / len(marks)
    else:
        avg = None
    student_info["Average"] = avg

    # Extra text
    for col in text_cols:
        if col not in ["Grade"] and pd.notna(row[col]) and row[col] != "":
            student_info[col] = row[col]

    return student_info, None


def find_closest_name(name, df):
    names = df['Name'].tolist()
    matches = difflib.get_close_matches(name, names, n=1, cutoff=0.5)
    return matches[0] if matches else None


# ---------------- PDF EXPORT -------------------------
def export_pdf(student_info, local_summary, gemini_summary, improvements_summary, prediction_summary, filename="AttendSmart_Report.pdf"):
    doc = SimpleDocTemplate(filename, pagesize=A4)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="Title",
        parent=styles["Heading1"],
        fontSize=18,
        textColor=colors.HexColor("#2E86C1"),
        alignment=1,
    )
    section_style = ParagraphStyle(
        name="Section",
        parent=styles["Heading2"],
        fontSize=14,
        textColor=colors.HexColor("#1B4F72"),
    )
    normal_style = styles["Normal"]

    # Dual Logos - AttendSmart and School
    left_logo = None
    right_logo = None
    
    # Left logo (AttendSmart)
    try:
        left_logo = Image("icon.ico", width=80, height=80)
    except Exception:
        left_logo = Paragraph("AttendSmart", title_style)
    
    # Right logo (School)
    try:
        right_logo = Image("school_icon.jpg", width=80, height=80)
    except Exception:
        right_logo = Paragraph("School", title_style)
    
    # Create table with both logos
    logo_table = Table([
        [left_logo, right_logo]
    ], colWidths=[200, 200])
    
    logo_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    
    story.append(logo_table)

    story.append(Spacer(1, 12))
    story.append(Paragraph("📖 Student Progress Report", title_style))
    story.append(Spacer(1, 20))

    # Student Info
    story.append(Paragraph(f"<b>Name:</b> {student_info.get('Name','-')}", normal_style))
    story.append(Paragraph(f"<b>Grade:</b> {student_info.get('Grade','-')}", normal_style))
    story.append(Spacer(1, 15))

    # Marks Table
    marks_data = [["Subject/Term", "Score"]]
    for key, value in student_info.items():
        if key not in ["Name", "Grade", "Average"]:
            marks_data.append([key, value])
    if "Average" in student_info and student_info["Average"] is not None:
        marks_data.append(["Average", f"{student_info['Average']:.1f}"])

    table = Table(marks_data, colWidths=[200, 200])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E86C1")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
    ]))
    story.append(table)
    story.append(Spacer(1, 20))

    # Excel Summary (Raw)
    story.append(Paragraph("📊 Basic Data", section_style))
    for line in local_summary.split("\n"):
        story.append(Paragraph(line, normal_style))
    story.append(Spacer(1, 15))

    # Gemini AI Summary
    story.append(Paragraph("🤖 AttendSmart AI Summary", section_style))
    for line in gemini_summary.split("\n"):
        story.append(Paragraph(line, normal_style))
    story.append(Spacer(1, 20))

    # Improvement Suggestions
    if improvements_summary and improvements_summary.strip():
        story.append(Paragraph("📌 Recommended Improvements", section_style))
        for line in improvements_summary.split("\n"):
            story.append(Paragraph(line, normal_style))
        story.append(Spacer(1, 20))

    # Future Outlook / Prediction
    if prediction_summary and prediction_summary.strip():
        story.append(Paragraph("🔮 Future Outlook", section_style))
        for line in prediction_summary.split("\n"):
            story.append(Paragraph(line, normal_style))
        story.append(Spacer(1, 20))

    # Footer
    footer = Paragraph("Generated by AttendSmart AI © 2025",
                       ParagraphStyle(name="Footer", fontSize=8,
                                      textColor=colors.grey, alignment=1))
    story.append(footer)

    doc.build(story)
    return filename


# ---------------- ACTIONS ----------------------------
def generate_summary():
    student_name = name_entry.get().strip()
    api_key = api_entry.get().strip()

    if not student_name:
        messagebox.showwarning("Input Error", "Please enter a student name.")
        return
    if not api_key:
        messagebox.showwarning("Input Error", "Please enter Gemini API key.")
        return

    closest_name = find_closest_name(student_name, marks_df)
    if not closest_name:
        messagebox.showwarning("Not Found", f"No record found for '{student_name}'.")
        return

    student_info, error = summarize_student(closest_name, marks_df)
    if error:
        messagebox.showwarning("Error", error)
        return

    # Local summary text
    local_summary = f"Student: {student_info['Name']} (Grade {student_info['Grade']})\n"
    for k, v in student_info.items():
        if k not in ["Name", "Grade"]:
            local_summary += f"{k}: {v}\n"

    local_textbox.delete("1.0", "end")
    local_textbox.insert("end", local_summary.strip())

    try:
        gemini_client = genai.Client(api_key=api_key)
        response = gemini_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=f"Summarize the following student data for a school report for the principal embed your idea within it too and only give these information and dont use emoji's and dont use points:\n{local_summary}"
        )
        gemini_textbox.delete("1.0", "end")
        gemini_textbox.insert("end", response.text)

        # Generate improvements advice
        improvements = gemini_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=(
                "Given the student's data below, provide concise, actionable improvement suggestions "
                "for the next term. Use plain sentences, no bullet points, no emojis. 4-6 sentences max.\n" 
                + local_summary
            )
        )
        improvements_textbox.delete("1.0", "end")
        improvements_textbox.insert("end", improvements.text)

        # Generate future prediction / outlook
        prediction = gemini_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=(
                "Based on the student's current performance, attendance and conduct, "
                "write a realistic future outlook for the next academic year and medium term. "
                "Avoid lists and emojis; keep it 4-6 sentences, plain sentences.\n" 
                + local_summary
            )
        )
        prediction_textbox.delete("1.0", "end")
        prediction_textbox.insert("end", prediction.text)
    except Exception as e:
        messagebox.showerror("AttendSmart AI Error", str(e))


def save_report():
    student_name = name_entry.get().strip()
    local_summary = local_textbox.get("1.0", "end").strip()
    gemini_summary = gemini_textbox.get("1.0", "end").strip()
    improvements_summary = improvements_textbox.get("1.0", "end").strip()
    prediction_summary = prediction_textbox.get("1.0", "end").strip()

    if not student_name or not local_summary:
        messagebox.showwarning("Error", "Please generate a summary first.")
        return

    # Get student_info safely
    closest_name = find_closest_name(student_name, marks_df)
    if not closest_name:
        messagebox.showwarning("Error", f"No record found for '{student_name}'.")
        return

    student_info, error = summarize_student(closest_name, marks_df)
    if error or not student_info:
        messagebox.showwarning("Error", "Student info not available. Cannot save PDF.")
        return

    filename = asksaveasfilename(defaultextension=".pdf",
                                 filetypes=[("PDF Files", "*.pdf")])
    if not filename:
        return

    try:
        export_pdf(student_info, local_summary, gemini_summary, improvements_summary, prediction_summary, filename)
        messagebox.showinfo("Success", f"Report saved as {filename}")
    except Exception as e:
        messagebox.showerror("Error", f"Could not save PDF:\n{e}")


def change_theme(choice):
    ctk.set_appearance_mode(choice)


# ---------------- MODERN TKINTER (CTK) ----------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.title("AttendSmart AI Digest")
root.geometry("850x700")
root.minsize(500, 400)

# Header
header = ctk.CTkFrame(root, corner_radius=0)
header.pack(fill="x")

title_label = ctk.CTkLabel(header,
                           text="📊 AttendSmart AI Digest Principal Dashboard",
                           font=("Helvetica", 18, "bold"))
title_label.pack(side="left", padx=10, pady=10)

theme_opt = ctk.CTkOptionMenu(header,
                              values=["Dark", "Light", "System"],
                              command=change_theme)
theme_opt.set("Dark")
theme_opt.pack(side="right", padx=10, pady=10)

# Input Frame
frame = ctk.CTkFrame(root)
frame.pack(padx=20, pady=10, fill="x")

ctk.CTkLabel(frame, text="Student Name:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
name_entry = ctk.CTkEntry(frame, width=300)
name_entry.grid(row=0, column=1, padx=5, pady=5)

ctk.CTkLabel(frame, text="Gemini API Key:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
api_entry = ctk.CTkEntry(frame, width=300, show="*")
api_entry.grid(row=1, column=1, padx=5, pady=5)

generate_btn = ctk.CTkButton(frame,
                             text="Generate Summary",
                             command=generate_summary,
                             fg_color="#4caf50",
                             hover_color="#45a049")
generate_btn.grid(row=2, column=0, columnspan=2, pady=10)

save_btn = ctk.CTkButton(frame,
                         text="💾 Save Report as PDF",
                         command=save_report,
                         fg_color="#2196f3",
                         hover_color="#1976d2")
save_btn.grid(row=3, column=0, columnspan=2, pady=10)

# Summary Boxes
ctk.CTkLabel(root, text="Excel Summary:", font=("Helvetica", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 2))
local_textbox = ctk.CTkTextbox(root, height=10, corner_radius=8)
local_textbox.pack(padx=20, pady=5, fill="both", expand=True)

ctk.CTkLabel(root, text="AttendSmart AI Summary:", font=("Helvetica", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 2))
gemini_textbox = ctk.CTkTextbox(root, height=10, corner_radius=8)
gemini_textbox.pack(padx=20, pady=5, fill="both", expand=True)

# Improvements Section
ctk.CTkLabel(root, text="Recommended Improvements:", font=("Helvetica", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 2))
improvements_textbox = ctk.CTkTextbox(root, height=8, corner_radius=8)
improvements_textbox.pack(padx=20, pady=5, fill="both", expand=True)

# Prediction Section
ctk.CTkLabel(root, text="Future Outlook:", font=("Helvetica", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 2))
prediction_textbox = ctk.CTkTextbox(root, height=8, corner_radius=8)
prediction_textbox.pack(padx=20, pady=5, fill="both", expand=True)

root.mainloop()
