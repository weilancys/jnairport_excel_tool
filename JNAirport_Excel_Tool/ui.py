import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox
from .excel import find_rows_30days_from_now

class JNAirport_Excel_Tool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.__init_ui()
    
    
    def __init_ui(self):
        self.title("济南机场 EXCEL 工具")
        self.geometry("450x200")

        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True, fill=tk.BOTH)

        description = ttk.Label(main_frame)
        description["text"] = """
        选择相应的Excel文件，自动筛选出截至缴费时间至今不足30天的记录。
        筛选出的记录将自动通过邮件发送。
        """
        description.pack(expand=True, fill=tk.BOTH)

        btn_choose_excel_file = ttk.Button(main_frame, text="选择Excel文件", command=self.on_btn_choose_excel_file_click)
        btn_choose_excel_file.pack(pady=30)

    
    def on_btn_choose_excel_file_click(self):
        excel_filename = tk.filedialog.askopenfilename()
        if excel_filename == "":
            return
        if not excel_filename.endswith(".xlsx"):
            tk.messagebox.showerror("错误", "非Excel文件")
            return
        
        rows = find_rows_30days_from_now(excel_filename)
        if rows == []:
            tk.messagebox.showinfo("筛选结果", "没有发现 截至缴费时间 至今 不足30天的记录")
            return
        
        report_txt = "筛选出 {row_count} 条截至缴费时间 至今 不足30天的记录\n是否将以下结果通过邮件发送？\n".format(row_count=len(rows))
        for row in rows:
            report_txt += "\n{company_name} - {contact} - {phone} - {deadline}".format(
                company_name = row[1],
                contact = row[4],
                phone = row[5],
                deadline = row[19]
            )

        need_email_sent = tk.messagebox.askyesno("筛选结果", report_txt)
        if need_email_sent:
            pass


    def run(self):
        self.mainloop()


if __name__ == "__main__":
    tool = JNAirport_Excel_Tool()
    tool.run()