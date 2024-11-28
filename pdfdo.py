import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter
import io
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import pandas as pd

# 注册中文字体
pdfmetrics.registerFont(TTFont("SimSun", "simsun.ttc"))  # 替换为系统中的字体路径

# 创建文字叠加 PDF 页面
def create_overlay_page(text, x, y, font_size):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("SimSun", font_size)
    can.drawString(x, y, text)
    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0]

# 单页插入文字
def add_text_to_pdf(input_pdf, output_pdf, text, page_number, x, y, font_size):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        if i == page_number - 1:
            overlay = create_overlay_page(text, x, y, font_size)
            page.merge_page(overlay)  # 合并文字到指定页
        writer.add_page(page)

    with open(output_pdf, "wb") as output_stream:
        writer.write(output_stream)

# 使用 CSV 文件生成多页 PDF
def replicate_pdf_with_text(input_pdf, output_pdf, csv_file, page_number, x, y, font_size):
    df = pd.read_csv(csv_file)
    if df.empty or df.shape[1] < 1:
        raise ValueError("CSV 文件为空或缺少列")

    text_list = df.iloc[:, 0].tolist()
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for text in text_list:
        overlay = create_overlay_page(text, x, y, font_size)
        base_page = reader.pages[page_number - 1]
        combined_page = PageObject.create_blank_page(width=base_page.mediabox.width, height=base_page.mediabox.height)
        combined_page.merge_page(base_page)
        combined_page.merge_page(overlay)
        writer.add_page(combined_page)

    with open(output_pdf, "wb") as output_stream:
        writer.write(output_stream)

# PDF 预览功能
def render_pdf_to_image(pdf_path, page_number):
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_number - 1)
        pix = page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img
    except Exception as e:
        messagebox.showerror("错误", f"预览PDF时出错：{e}")
        return None

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_path_var.set(file_path)

def select_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    csv_path_var.set(file_path)

def save_pdf():
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    save_path_var.set(file_path)

def preview_text():
    try:
        page_num = int(page_number_var.get())
        x = int(x_var.get())
        y = int(y_var.get())
        text = text_var.get()
        font_size = int(font_size_var.get())
        input_pdf = pdf_path_var.get()

        if not input_pdf or not text:
            messagebox.showerror("错误", "请填写所有字段并选择文件")
            return

        preview_pdf = "preview_temp.pdf"
        add_text_to_pdf(input_pdf, preview_pdf, text, page_num, x, y, font_size)

        img = render_pdf_to_image(preview_pdf, page_num)
        if img:
            original_width, original_height = img.size
            max_width, max_height = 600, 800
            scale = min(max_width / original_width, max_height / original_height)
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

            img_tk = ImageTk.PhotoImage(img)
            preview_canvas.delete("all")
            preview_canvas.create_image((max_width - new_width) // 2, (max_height - new_height) // 2, anchor="nw", image=img_tk)
            preview_canvas.image = img_tk

    except ValueError:
        messagebox.showerror("错误", "请确保页码、X和Y坐标、字体大小为有效数字")

def apply_text_to_pdf():
    try:
        page_num = int(page_number_var.get())
        x = int(x_var.get())
        y = int(y_var.get())
        text = text_var.get()
        font_size = int(font_size_var.get())
        input_pdf = pdf_path_var.get()
        output_pdf = save_path_var.get()

        if not input_pdf or not output_pdf or not text:
            messagebox.showerror("错误", "请填写所有字段并选择文件")
            return

        add_text_to_pdf(input_pdf, output_pdf, text, page_num, x, y, font_size)
        messagebox.showinfo("成功", f"文字已插入到PDF并保存为 {output_pdf}")
    except Exception as e:
        messagebox.showerror("错误", f"处理PDF时出错：{e}")

def apply_csv_to_replicated_pdf():
    try:
        page_num = int(page_number_var.get())
        x = int(x_var.get())
        y = int(y_var.get())
        font_size = int(font_size_var.get())
        input_pdf = pdf_path_var.get()
        output_pdf = save_path_var.get()
        csv_file = csv_path_var.get()

        if not input_pdf or not output_pdf or not csv_file:
            messagebox.showerror("错误", "请填写所有字段并选择文件")
            return

        replicate_pdf_with_text(input_pdf, output_pdf, csv_file, page_num, x, y, font_size)
        messagebox.showinfo("成功", f"CSV 数据已插入到PDF并保存为 {output_pdf}")
    except Exception as e:
        messagebox.showerror("错误", f"处理CSV时出错：{e}")

# 创建主窗口
root = tk.Tk()
root.title("PDF文字插入器")

# 输入字段
tk.Label(root, text="PDF路径:").grid(row=0, column=0, sticky="e")
pdf_path_var = tk.StringVar()
tk.Entry(root, textvariable=pdf_path_var, width=40).grid(row=0, column=1)
tk.Button(root, text="选择文件", command=select_pdf).grid(row=0, column=2)

tk.Label(root, text="CSV路径:").grid(row=1, column=0, sticky="e")
csv_path_var = tk.StringVar()
tk.Entry(root, textvariable=csv_path_var, width=40).grid(row=1, column=1)
tk.Button(root, text="选择文件", command=select_csv).grid(row=1, column=2)

tk.Label(root, text="保存路径:").grid(row=2, column=0, sticky="e")
save_path_var = tk.StringVar()
tk.Entry(root, textvariable=save_path_var, width=40).grid(row=2, column=1)
tk.Button(root, text="保存文件", command=save_pdf).grid(row=2, column=2)

tk.Label(root, text="页码:").grid(row=3, column=0, sticky="e")
page_number_var = tk.StringVar()
tk.Entry(root, textvariable=page_number_var, width=10).grid(row=3, column=1, sticky="w")

tk.Label(root, text="插入文字:").grid(row=4, column=0, sticky="e")
text_var = tk.StringVar()
tk.Entry(root, textvariable=text_var, width=40).grid(row=4, column=1, columnspan=2)

tk.Label(root, text="X坐标:").grid(row=5, column=0, sticky="e")
x_var = tk.StringVar()
tk.Entry(root, textvariable=x_var, width=10).grid(row=5, column=1, sticky="w")

tk.Label(root, text="Y坐标:").grid(row=6, column=0, sticky="e")
y_var = tk.StringVar()
tk.Entry(root, textvariable=y_var, width=10).grid(row=6, column=1, sticky="w")

tk.Label(root, text="字体大小:").grid(row=7, column=0, sticky="e")
font_size_var = tk.StringVar(value="12")
tk.Entry(root, textvariable=font_size_var, width=10).grid(row=7, column=1, sticky="w")

# 按钮和画布
tk.Button(root, text="预览", command=preview_text).grid(row=8, column=0, columnspan=3, pady=5)
tk.Button(root, text="应用文字到PDF", command=apply_text_to_pdf).grid(row=9, column=0, columnspan=3, pady=5)
tk.Button(root, text="应用CSV到复制PDF", command=apply_csv_to_replicated_pdf).grid(row=10, column=0, columnspan=3, pady=5)

preview_canvas = tk.Canvas(root, width=600, height=800, bg="white")
preview_canvas.grid(row=11, column=0, columnspan=3)

root.mainloop()
