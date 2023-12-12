import tkinter as tk 
from tkinter import ttk, messagebox, filedialog, simpledialog 
import TKinterModernThemes as TKMT 
import re 
import docx 
import openpyxl 
from PyPDF2 import PdfReader 
import csv 
import requests 
from bs4 import BeautifulSoup 
import os 
import webbrowser 
import pyperclip

def createToolTip(widget, text):
    try:
        def enter(event):
            widget._after_id = widget.after(600, show_tooltip, event)

        def leave(event):
            widget.after_cancel(widget._after_id)
            tooltip = getattr(widget, "_tooltip", None)
            if tooltip:
                tooltip.destroy()
                widget._tooltip = None

        def show_tooltip(event):
            tooltip = tk.Toplevel(widget)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root}+{event.y_root}")
            label = tk.Label(tooltip, text=text, background="black", foreground="white")
            label.grid()
            widget._tooltip = tooltip

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

class App(TKMT.ThemedTKinterFrame):
    def __init__(self):
        try:
            super().__init__("Word Counter Master", "Sun-valley", "dark")
            self.master.iconbitmap('Word Counter Master.ico')
            self.master.resizable(False, False)

            # 创建一个主框架，用于包含所有的小部件
            self.main_frame = ttk.Frame(self.master)
            self.main_frame.grid(row=0, column=0, sticky="nsew")

            # 创建一个左侧框架，用于显示工具按钮
            self.left_frame = ttk.Frame(self.main_frame, width=200, height=400)
            self.left_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ns")

            # 创建一个右侧框架，用于显示文本输入和计数标签
            self.right_frame = ttk.Frame(self.main_frame)
            self.right_frame.grid(row=0, column=1, padx=10, pady=5, sticky="nsew")

            # 在右侧框架中创建一个文本输入框
            self.text_input = tk.Text(self.right_frame, font=("Arial", 12))
            self.text_input.grid(row=0, column=0, sticky="nsew")
            createToolTip(self.text_input, "Enter the text you want to count here")
            self.text_input.bind("<KeyRelease>", self.count_words)

            # 在右侧框架中创建一个滚动条，与文本输入框关联
            self.scrollbar = ttk.Scrollbar(self.right_frame, orient="vertical", command=self.text_input.yview)
            self.scrollbar.grid(row=0, column=1, sticky="ns")
            self.text_input.config(yscrollcommand=self.scrollbar.set)

            # 在右侧框架中创建一个计数标签，用于显示各种计数
            self.label_count = ttk.Label(self.right_frame, text="", font=("Arial", 10))
            createToolTip(self.label_count, "This displays the word count")
            self.label_count.grid(row=1, column=0, sticky="ns")
            self.count_words()

            # 在左侧框架中创建一个打开文件按钮
            self.open_button = ttk.Button(self.left_frame, text="Open File", command=self.open_file)
            self.open_button.grid(row=0, column=0, sticky="ew", pady=10)
            createToolTip(self.open_button, "Click to open a file")

            # 在左侧框架中创建一个保存文本按钮
            self.save_button = ttk.Button(self.left_frame, text="Save Text", command=self.save_file)
            self.save_button.grid(row=1, column=0, sticky="ew", pady=10)
            createToolTip(self.save_button, "Click to save the text")

            # 在左侧框架中创建一个复制文本按钮
            self.copy_button = ttk.Button(self.left_frame, text="Copy Text", command=self.copy_text)
            self.copy_button.grid(row=2, column=0, sticky="ew", pady=10)
            createToolTip(self.copy_button, "Click to copy the text")

            # 在左侧框架中创建一个清空并复制到剪贴板按钮
            self.clear_copy_button = ttk.Button(self.left_frame, text="Clear and Copy to Clipboard", command=self.clear_and_copy_to_clipboard)
            self.clear_copy_button.grid(row=3, column=0, sticky="ew", pady=10)
            createToolTip(self.clear_copy_button, "Clear the text box and copy its contents to clipboard")

            # 在左侧框架中创建一个清空文本按钮
            self.clear_button = ttk.Button(self.left_frame, text="Clear Text", command=self.clear_text)
            self.clear_button.grid(row=4, column=0, sticky="ew", pady=10)
            createToolTip(self.clear_button, "Clear the text box")

            # 在左侧框架中创建一个获取网页内容按钮
            self.fetch_button = ttk.Button(self.left_frame, text="Fetch Web Content", command=self.fetch_web_content)
            self.fetch_button.grid(row=5, column=0, sticky="ew", pady=10)
            createToolTip(self.fetch_button, "Click to fetch web content")

            # 在左侧框架中创建一个发送反馈按钮
            self.feedback_button = ttk.Button(self.left_frame, text="Send Feedback", command=self.open_feedback_link)
            self.feedback_button.grid(row=6, column=0, sticky="ew", pady=10)
            createToolTip(self.feedback_button, "Click to send feedback")

            # 在左侧框架中创建一个查看快捷键按钮
            self.shortcuts_button = ttk.Button(self.left_frame, text="View Shortcuts", command=self.view_shortcuts)
            self.shortcuts_button.grid(row=7, column=0, sticky="ew", pady=10)
            createToolTip(self.shortcuts_button, "Click to view all shortcuts")

            # 设置主框架的网格权重，使其能够自适应窗口大小
            self.main_frame.columnconfigure(1, weight=1)
            self.main_frame.rowconfigure(0, weight=1)

            # 设置右侧框架的网格权重，使其能够自适应窗口大小
            self.right_frame.columnconfigure(0, weight=1)
            self.right_frame.rowconfigure(0, weight=1)

            # 绑定一些快捷键到窗口
            self.master.bind('<Control-o>', lambda event: self.open_file())
            self.master.bind('<Control-s>', lambda event: self.save_file())
            self.master.bind('<Control-c>', lambda event: self.copy_text())
            self.master.bind('<Control-x>', lambda event: self.clear_text())
            self.master.bind('<Control-f>', lambda event: self.fetch_web_content())
            self.master.bind('<Control-e>', lambda event: self.open_feedback_link())

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def open_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[('Text Documents', '*.txt *.docx *.doc *.xlsx *.xls *.pdf *.csv')])
            if file_path:
                if file_path.endswith('.txt'):
                    with open(file_path, 'r') as f:
                        text = f.read()

                elif file_path.endswith('.doc') or file_path.endswith('.docx'):
                    doc = docx.Document(file_path)
                    text = '\n'.join([para.text for para in doc.paragraphs])

                elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                    wb = openpyxl.load_workbook(file_path)
                    text = ""
                    for sheet in wb:
                        for row in sheet.rows:
                            for cell in row:
                                if isinstance(cell.value, str):
                                    text += cell.value
                                else:
                                    text += str(cell.value)

                elif file_path.endswith('.pdf'):
                    pdfFileObj = open(file_path, 'rb')
                    pdfReader = PdfReader(pdfFileObj)
                    text = ""
                    for pageNum in range(len(pdfReader.pages)):
                        pageObj = pdfReader.pages[pageNum]
                        text += pageObj.extract_text()
                    pdfFileObj.close()

                elif file_path.endswith('.csv'):
                    with open(file_path, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        text = '\n'.join([' '.join(row) for row in reader])
                    
                else:
                    return
                    
                self.text_input.delete('1.0', tk.END)
                self.text_input.insert(tk.END, text)
                self.count_words()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def fetch_web_content(self):
        url = simpledialog.askstring("Input", "Enter the URL")
        if url:
            try:
                response = requests.get(url)
                soup = BeautifulSoup(response.text, 'html.parser')
                text = soup.get_text()
                self.text_input.delete('1.0', tk.END)
                self.text_input.insert(tk.END, text)
                self.count_words()
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def count_words(self, event=None):
        try:
            text = self.text_input.get("1.0", "end")
            word_counts = self.count_all_words(text)
            self.label_count.config(text=word_counts)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def count_all_words(self, text):
        try:
            counts = {
                "Chinese Count": len(re.findall(r'[\u4E00-\u9FFF]', text)),
                "English Count": len(re.findall(r'\b[a-zA-Z]+\b', text)),
                "Malay Count": len(re.findall(r'\b[abcçdefgğhıijklmnoöprsştuüvyz]+([\'-][abcçdefgğhıijklmnoöprsştuüvyz]+)*\b', text, re.IGNORECASE)),
                "Number Count": len(re.findall(r'\d+', text)),
                "Comma Count": len(re.findall(r',|，', text)),
                "Period Count": len(re.findall(r'\.|。', text)),
            }
            return "\n".join(f"{key}: {value}" for key, value in counts.items())
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def save_file(self):
        try:
            text = self.text_input.get("1.0", "end").strip()
            if not text:
                messagebox.showwarning("Info", "Input Text In The Text Box First!")
                return

            file_name = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[('Text Documents', '*.txt'), ('Word Documents', '*.docx'), ('PDF Files', '*.pdf')], initialfile='textbox')
            if file_name:
                extension = os.path.splitext(file_name)[1].lower()

                if extension == '.txt':
                    with open(file_name, 'w', encoding='utf-8') as f:
                        f.write(text)

                elif extension == '.docx':
                    doc = docx.Document()
                    doc.add_paragraph(text)
                    doc.save(file_name)

                elif extension == '.pdf':
                    from reportlab.platypus import SimpleDocTemplate, Paragraph
                    from reportlab.lib.styles import getSampleStyleSheet
                    from reportlab.lib.pagesizes import letter
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont

                    pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))  
                    pdf = SimpleDocTemplate(file_name, pagesize=letter)
                    styles = getSampleStyleSheet()
                    style = styles['Normal']
                    style.fontName = 'SimSun'  
                    style.fontSize = 15
                    story = [Paragraph(text, style)]
                    pdf.build(story)

                else:
                    messagebox.showwarning("Warning", "Unsupported file format")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def open_feedback_link(self):
        webbrowser.open("https://github.com/fatherxtreme123/Word.Counter.Master/issues")

    def copy_text(self):
        try:
            text = self.text_input.get("1.0", "end").strip()
            if text:
                pyperclip.copy(text)
            else:
                messagebox.showinfo("Info", "Text box is empty.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def clear_and_copy_to_clipboard(self):
        try:
            text = self.text_input.get("1.0", "end").strip()
            if text:
                self.text_input.delete('1.0', tk.END)
                pyperclip.copy(text)
            else:
                messagebox.showinfo("Info", "Text box is already empty.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def clear_text(self):
        try:
            if self.text_input.get("1.0", "end").strip():
                self.text_input.delete('1.0', tk.END)
                self.count_words()
            else:
                messagebox.showinfo("Info", "Text box is already empty.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def view_shortcuts(self):
        try:
            shortcuts_info = """
            Shortcuts:
            - Open File: Ctrl + O
            - Save Text: Ctrl + S
            - Copy Text: Ctrl + C
            - Clear Text: Ctrl + X
            - Fetch Web Content: Ctrl + F
            - Send Feedback: Ctrl + E
            """

            shortcuts_window = tk.Toplevel(self.master)
            shortcuts_window.title("Shortcuts")
            shortcuts_window.resizable(False, False)

            label = ttk.Label(shortcuts_window, text=shortcuts_info)
            label.grid(row=0, column=0, padx=10, pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    try:
        App().run()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")