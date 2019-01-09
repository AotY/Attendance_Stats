import queue
import platform
import tkinter as tk
from tkinter import filedialog
from attd_record import AttdRecord
import util
import os


class GradeGui(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("教务处考勤统计汇总")
        util.center_window(self, 800, 500)
        self.resizable(False, False)

        standard_font = (None, 14)

        self.main_frame = tk.Frame(self, width=550, height=350, bg="lightgrey")

        self.button_frame = tk.Frame(self, width=500, height=100)

        self.stats_btn = tk.Button(self.button_frame, text="开始统计", bg="lightgrey", fg="black",
                                      command=self.stats, font=standard_font, state="active")
        self.stats_btn.pack(side=tk.LEFT)

        self.button_frame.pack(side=tk.BOTTOM, pady=10)

        self.parameters_frame = tk.Frame(self, width=220, height=200)
        self.parameters_frame.pack(side=tk.LEFT, padx=3)

        # 上午打卡时间
        am_frame = tk.Frame(self.parameters_frame, width=200, height=50)
        am_label = tk.Label(am_frame, text='上午打卡时间：', font=standard_font)
        am_label.pack(side=tk.LEFT, padx=2)

        self.am_var = tk.StringVar()
        self.am_entry = tk.Entry(am_frame, textvariable=self.am_var, font=standard_font)
        self.am_entry.pack(side=tk.RIGHT, padx=2)
        self.am_default_v = '默认：08:00'
        self.am_var.set(self.am_default_v)
        am_frame.pack(side=tk.TOP, pady=5)

        # 下午打卡时间
        pm_frame = tk.Frame(self.parameters_frame, width=200, height=50)
        pm_label = tk.Label(pm_frame, text='下午打卡时间：', font=standard_font)
        pm_label.pack(side=tk.LEFT, padx=2)

        self.pm_var = tk.StringVar()
        self.pm_entry = tk.Entry(pm_frame, textvariable=self.pm_var, font=standard_font)
        self.pm_entry.pack(side=tk.RIGHT, padx=2)
        self.pm_default_v = '默认：14:00'
        self.pm_var.set(self.pm_default_v)
        pm_frame.pack(side=tk.TOP, pady=5)

        # 打卡时间表
        schedule_frame = tk.Frame(self.parameters_frame, width=200, height=50)
        schedule_label = tk.Label(schedule_frame, text='打卡时间表：', font=standard_font)
        schedule_label.pack(side=tk.LEFT, padx=2)

        self.schedule_var = tk.StringVar()
        self.schedule_entry = tk.Entry(schedule_frame, textvariable=self.schedule_var)
        self.schedule_entry.pack(side=tk.RIGHT, padx=2)

        self.schedule_btn = tk.Button(schedule_frame, text="选择文件", bg="lightgrey", fg="black",
                                      command=lambda: self.choose_file(1), font=standard_font, state="active")
        self.schedule_btn.pack(side=tk.RIGHT, padx=2)

        schedule_frame.pack(side=tk.TOP, pady=5)

        # 月度汇总表
        summary_frame = tk.Frame(self.parameters_frame, width=200, height=50)
        summary_label = tk.Label(summary_frame, text='月度汇总表：', font=standard_font)
        summary_label.pack(side=tk.LEFT, padx=2)

        self.summary_var = tk.StringVar()
        self.summary_entry = tk.Entry(summary_frame, textvariable=self.summary_var)
        self.summary_entry.pack(side=tk.RIGHT, padx=2)

        self.summary_btn = tk.Button(summary_frame, text="选择文件", bg="lightgrey", fg="black",
                                     command=lambda: self.choose_file(2), font=standard_font, state="active")
        self.summary_btn.pack(side=tk.RIGHT, padx=2)
        summary_frame.pack(side=tk.TOP, pady=5)

        # 保存路径
        save_frame = tk.Frame(self.parameters_frame, width=200, height=50)
        save_label = tk.Label(save_frame, text='保存目录：', font=standard_font)
        save_label.pack(side=tk.LEFT, padx=2)

        self.save_var = tk.StringVar()
        self.save_entry = tk.Entry(save_frame, textvariable=self.save_var, font=standard_font)
        self.save_entry.pack(side=tk.RIGHT, padx=2)

        self.choose_dir_btn = tk.Button(save_frame, text="选择目录", bg="lightgrey", fg="black",
                                      command=self.choose_dir, font=standard_font, state="active")
        self.choose_dir_btn.pack(side=tk.RIGHT, padx=2)

        save_frame.pack(side=tk.TOP, pady=5)

        # 输出
        self.output_frame = tk.Frame(self, width=220, height=200)
        self.output_frame.pack(side=tk.RIGHT, padx=5)

        self.output_text = tk.Text(self.output_frame, bg="#f2f2f2", fg="black", font=standard_font)
        self.output_text.pack(side=tk.TOP, expand=1, pady=10)

        self.output_text.delete('1.0', tk.END)
        self.output_text['state'] = tk.DISABLED

        self.main_frame.pack(fill=tk.BOTH, expand=1)
        self.protocol("WM_DELETE_WINDOW", self.safe_destroy)

    # 选择文件
    def choose_file(self, type=1):
        self.output_text['state'] = tk.NORMAL

        if (platform.system() == 'Windows'):
            file_path = filedialog.askopenfilename(title="选择文件",
                                                        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        else:
            file_path = filedialog.askopenfilename(title="选择文件",
                                                        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if os.path.exists(file_path):
            print(file_path)
            if type == 1:
                self.schedule_var.set(file_path)
                self.output_text.insert(tk.INSERT, "打卡时间表: {} \n".format(file_path))
            elif type == 2:
                self.summary_var.set(file_path)
                self.output_text.insert(tk.INSERT, "月度汇总表: {} \n".format(file_path))
        else:
            print("文件不存在或打开错误")
            self.output_text.insert(tk.INSERT, "文件不存在或打开错误，请重新打开文件 \n")

    # choose dir
    def choose_dir(self):
        self.output_text['state'] = tk.NORMAL
        save_dir = filedialog.askdirectory()

        if os.path.exists(save_dir):
            print('save_dir: %s' % save_dir)
        else:
            self.output_text.insert(tk.INSERT, "目录不存在，请重新选择 \n")
        self.save_var.set(save_dir)

    def stats(self):
        time_am = self.am_var.get()
        if time_am == '' or time_am == self.am_default_v:
            time_am = '08:00'
        self.output_text.insert(tk.INSERT, "上午打卡时间：%s\n"  % time_am)

        time_pm = self.pm_var.get()
        if time_pm == '' or time_pm == self.pm_default_v:
            time_pm = '14:00'
        self.output_text.insert(tk.INSERT, "下午打卡时间：%s\n"  % time_pm)

        schedule_path = ''
        schedule_path = self.schedule_var.get()
        if schedule_path == '' or not os.path.exists(schedule_path):
            self.output_text.insert(tk.INSERT, "打卡时间表不存在，请重新选择 \n")
            return

        summary_path = self.summary_var.get()
        if summary_path == '' or not os.path.exists(summary_path):
            self.output_text.insert(tk.INSERT, "月度汇总表不存在，请重新选择 \n")
            return

        save_dir = self.save_var.get()
        if save_dir == '' or not os.path.exists(save_dir):
            self.output_text.insert(tk.INSERT, "保存不存在，请重新选择 \n")
            return

        #  try:
        self.queue = queue.Queue()
        attd_record = AttdRecord(self.queue, schedule_path, summary_path,
                                 save_dir, time_am, time_pm)
        attd_record.start()

        self.after(100, self.process_queue)
        self.stats_btn['state'] = tk.DISABLED

        #  except Exception as e:
            #  self.output_text.insert(tk.INSERT, "统计出差，请检查设置 \n")
            #  return
        #  self.output_text.insert(tk.INSERT, "统计完成\n")

    def process_queue(self):
        try:
            msg = self.queue.get(0)
            self.output_text.insert(tk.INSERT, "%s\n" % msg)
            self.stats_btn['state'] = tk.NORMAL
        except queue.Empty:
            #  self.output_text.insert(tk.INSERT, "统计出差，请检查设置 \n")
            self.after(100, self.process_queue)

    def safe_destroy(self):
        self.destroy()


if __name__ == "__main__":
    gui = GradeGui()
    gui.mainloop()
