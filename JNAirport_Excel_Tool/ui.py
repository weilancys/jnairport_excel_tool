import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox
from .excel import find_noticeable_rows
from .email import prepare_email, send_notice_mail


class JNAirport_Excel_Tool(tk.Tk):
    def __init__(self, smtp_auth, recipients):
        super().__init__()
        self.__init_ui()
        self.smtp_auth = smtp_auth # smtp config dict
        self.recipients = recipients # recipients list

    
    def __init_ui(self):
        self.title("济南机场 EXCEL 工具")
        self.geometry("450x200")

        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True, fill=tk.BOTH)

        description = ttk.Label(main_frame)
        description["text"] = """
        使用说明：
        选择相应的Excel文件，自动筛选出缴费截止时间至今不足30天的记录。
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
        
        rows_lt_30_days, rows_expired = find_noticeable_rows(excel_filename)
        if rows_lt_30_days == [] and rows_expired == []:
            tk.messagebox.showinfo("筛选结果", "没有发现缴费截止时间至今不足30天或已过期的记录")
            return
        
        report_txt_for_lt_30_days = "筛选出 {row_count} 条缴费截止时间 至今 不足30天的记录:".format(row_count=len(rows_lt_30_days))
        for row in rows_lt_30_days:
            report_txt_for_lt_30_days += "\n{company_name} - {contact} - {phone} - {deadline}".format(
                company_name = row[1],
                contact = row[4],
                phone = row[5],
                deadline = row[19]
            )
        
        report_txt_for_expired = "筛选出 {row_count} 条已超期记录:".format(row_count=len(rows_expired))
        for row in rows_expired:
            report_txt_for_expired += "\n{company_name} - {contact} - {phone} - {deadline}".format(
                company_name = row[1],
                contact = row[4],
                phone = row[5],
                deadline = row[19]
            )

        report_txt = report_txt_for_lt_30_days + "\n\n" + report_txt_for_expired + "\n\n" + "是否将以下结果通过邮件发送？\n"
        
        need_email_sent = tk.messagebox.askyesno("筛选结果", report_txt)
        if need_email_sent:
            email_message = prepare_email(rows_lt_30_days, rows_expired, excel_filename, self.smtp_auth["SENDER_EMAIL"], self.recipients)
            try:
                send_notice_mail(email_message, self.recipients, self.smtp_auth)
                tk.messagebox.showinfo("筛选结果", "邮件已成功发送")
            except Exception as e:
                tk.messagebox.showerror("错误", "邮件发送失败\n" + str(e))


    def run(self):
        self.mainloop()
