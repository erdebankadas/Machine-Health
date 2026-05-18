import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors


df = None
machine_list = []


def clean_value(val):
    if pd.isna(val):
        return ""
    return str(val)


def upload_excel():
    global df, machine_list

    file = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")])
    if not file:
        return

    df = pd.read_excel(file, header=None)

    machine_list = []

    machine_type_row = df[df[1].astype(str).str.contains("Machine Type", case=False, na=False)].index[0]
    machine_no_row = df[df[1].astype(str).str.contains("Machine No", case=False, na=False)].index[0]

    for col in range(2, df.shape[1]):

        mtype = clean_value(df.iloc[machine_type_row, col])
        mno = clean_value(df.iloc[machine_no_row, col])

        if mtype and mno:
            machine_list.append((f"{mtype}_{mno}", col))

    machine_dropdown["values"] = [m[0] for m in machine_list]

    messagebox.showinfo("Success", "Excel Loaded Successfully")


def preview_machine():

    if df is None:
        return

    machine = machine_dropdown.get()

    machine_col = None
    for name, col in machine_list:
        if name == machine:
            machine_col = col
            break

    if machine_col is None:
        return

    preview = tk.Toplevel(root)
    preview.title("Machine Preview")
    preview.geometry("500x400")

    text = tk.Text(preview)
    text.pack(fill="both", expand=True)

    for i, row in df.iterrows():

        section = clean_value(row[0])
        parameter = clean_value(row[1])

        if parameter != "":
            value = ""
            if machine_col < len(row):
                value = clean_value(row[machine_col])

            text.insert("end", f"{section} | {parameter} | {value}\n")


def generate_pdf():

    if df is None:
        messagebox.showerror("Error", "Upload Excel first")
        return

    machine = machine_dropdown.get()

    machine_col = None
    for name, col in machine_list:
        if name == machine:
            machine_col = col
            break

    if machine_col is None:
        messagebox.showerror("Error", "Machine not found")
        return

    save_file = filedialog.asksaveasfilename(defaultextension=".pdf")

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'title',
        parent=styles['Heading1'],
        alignment=TA_CENTER,
        fontSize=16
    )

    center_style = ParagraphStyle(
        'center',
        parent=styles['Normal'],
        alignment=TA_CENTER,
        fontSize=10
    )

    cell_style = ParagraphStyle(
        'cell',
        parent=styles['Normal'],
        fontSize=7
    )

    elements = []

    elements.append(Paragraph("Shaktigarh Textiles And Industries Limited", title_style))
    elements.append(Paragraph("15/B, Hemanta Basu Sarani", center_style))
    elements.append(Paragraph("Kolkata – 700001", center_style))

    elements.append(Spacer(1, 8))

    elements.append(Paragraph("Machine Health Report", center_style))
    elements.append(Paragraph(f"Machine : {machine}", center_style))

    # -------- UNIT VALUE FROM FIRST ROW --------
    unit_value = ""
    if machine_col < df.shape[1]:
        unit_value = clean_value(df.iloc[0, machine_col])

    elements.append(Paragraph(f"Unit : {unit_value}", center_style))
    # ------------------------------------------

    elements.append(Spacer(1, 10))

    table_data = []
    spans = []

    current_section = ""
    start_row = 0
    row_index = 0

    # -------- IGNORE FIRST ROW --------
    for i, row in df.iloc[1:].iterrows():

        section = clean_value(row[0])
        parameter = clean_value(row[1])

        value = ""
        if machine_col < len(row):
            value = clean_value(row[machine_col])

        if section != "":
            if row_index > start_row:
                spans.append(('SPAN', (0, start_row), (0, row_index - 1)))

            current_section = section
            start_row = row_index

        if parameter != "":
            table_data.append([
                Paragraph(current_section, cell_style),
                Paragraph(parameter, cell_style),
                Paragraph(value, cell_style)
            ])
            row_index += 1

    spans.append(('SPAN', (0, start_row), (0, row_index - 1)))

    table = Table(
        table_data,
        colWidths=[180, 380, 120]
    )

    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
    ] + spans))

    elements.append(table)

    pdf = SimpleDocTemplate(
        save_file,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    pdf.build(elements)

    messagebox.showinfo("Success", "PDF Generated Successfully")

    root.destroy()


def exit_app():
    root.destroy()


# ---------------- GUI ---------------- #

root = tk.Tk()
root.title("Machine Health Report Generator")
root.geometry("520x380")
root.configure(bg="#1e3d59")


colors_cycle = ["#00ffff","#00ffcc","#00ff99","#00ff66","#00ffcc"]

def animate_title():
    current = top_title.cget("fg")
    next_color = colors_cycle[(colors_cycle.index(current)+1) % len(colors_cycle)]
    top_title.config(fg=next_color)
    root.after(300, animate_title)


top_title = tk.Label(
    root,
    text="Machine Health Check Up!!!!!",
    font=("Arial", 20, "bold"),
    bg="#1e3d59",
    fg="#00ffff"
)

top_title.pack(pady=5)

animate_title()


title = tk.Label(
    root,
    text="Machine Health Report Generator",
    font=("Arial", 16, "bold"),
    bg="#1e3d59",
    fg="white"
)
title.pack(pady=10)


glass = tk.Frame(
    root,
    bg="#2e5c78",
    bd=0
)

glass.place(relx=0.5, rely=0.5, anchor="center", width=420, height=220)


upload_btn = tk.Button(
    glass,
    text="Upload Excel",
    command=upload_excel,
    bg="#ff6e40",
    fg="white",
    font=("Arial", 11, "bold"),
    width=20
)
upload_btn.pack(pady=10)


tk.Label(
    glass,
    text="Select Machine",
    bg="#2e5c78",
    fg="white",
    font=("Arial", 11)
).pack()


machine_dropdown = ttk.Combobox(glass, width=40)
machine_dropdown.pack(pady=10)

machine_dropdown.bind("<<ComboboxSelected>>", lambda e: preview_machine())


generate_btn = tk.Button(
    glass,
    text="Generate PDF",
    command=generate_pdf,
    bg="#4caf50",
    fg="white",
    font=("Arial", 11, "bold"),
    width=20
)
generate_btn.pack(pady=10)


exit_btn = tk.Button(
    glass,
    text="Exit",
    command=exit_app,
    bg="#d32f2f",
    fg="white",
    font=("Arial", 11, "bold"),
    width=20
)
exit_btn.pack(pady=10)


credit = tk.Label(
    root,
    text="MADE BY DD",
    font=("Arial", 12, "bold"),
    bg="#1e3d59",
    fg="#ffd700"
)

credit.pack(side="bottom", pady=6)


root.mainloop()