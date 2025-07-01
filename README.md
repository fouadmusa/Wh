import pandas as pd
import time
import threading
import urllib.parse
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from tkinter import *
from tkinter import filedialog, messagebox, ttk, scrolledtext

# إعدادات
driver = None
is_paused = False
is_stopped = False
excel_data = None
attachment_paths = []
wait_time = 60
max_retries = 3

# استبدال متغيرات القالب في نص الرسالة
def render_template(text, row):
    for col in row.index:
        text = text.replace(f"{{{{{col}}}}}", str(row[col]))
    return text

# تحميل ملف Excel
def load_excel():
    global excel_data
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return
    try:
        excel_data = pd.read_excel(file_path)
        if not {"اسم العميل", "رقم العميل", "نص الرسالة"}.issubset(excel_data.columns):
            messagebox.showerror("خطأ", "الملف يجب أن يحتوي على الأعمدة: اسم العميل، رقم العميل، نص الرسالة")
            return
        if "تقرير الإرسال" not in excel_data.columns:
            excel_data["تقرير الإرسال"] = ""
        col_options = [col for col in excel_data.columns if col not in ["رقم العميل", "نص الرسالة", "تقرير الإرسال"]]
        filter_dropdown["values"] = col_options
        if col_options:
            filter_dropdown.current(0)
            update_filter_values()
        messagebox.showinfo("تم", f"تم تحميل {len(excel_data)} عميل.")
        log_event(f"تم تحميل ملف: {file_path}")
    except Exception as e:
        messagebox.showerror("خطأ", str(e))
        log_event(f"خطأ تحميل الملف: {str(e)}")

# اختيار مرفقات لإرسالها
def choose_attachments():
    global attachment_paths
    files = filedialog.askopenfilenames()
    if files:
        attachment_paths = list(files)
        messagebox.showinfo("📎 تم اختيار المرفقات", "\n".join(attachment_paths))
        log_event(f"تم اختيار المرفقات: {attachment_paths}")

# تحديث قيم الفلتر
def update_filter_values(*args):
    if excel_data is not None:
        col = filter_column.get()
        if col in excel_data.columns:
            values = excel_data[col].dropna().unique().tolist()
            filter_value_dropdown["values"] = values
            if values:
                filter_value_dropdown.current(0)

# البيانات بعد الفلترة
def get_filtered_data():
    if excel_data is None:
        return None
    col = filter_column.get()
    val = filter_value.get()
    if col in excel_data.columns:
        return excel_data[excel_data[col] == val].copy()
    return excel_data.copy()

# سجل الأحداث
def log_event(msg):
    log_text.config(state='normal')
    log_text.insert(END, f"{datetime.datetime.now().strftime('%H:%M:%S')} - {msg}\n")
    log_text.see(END)
    log_text.config(state='disabled')

# تصدير التقرير
def export_report():
    if excel_data is None:
        messagebox.showerror("خطأ", "لا يوجد تقرير للتصدير.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
    if not file_path:
        return
    try:
        if file_path.endswith(".csv"):
            excel_data.to_csv(file_path, index=False)
        else:
            excel_data.to_excel(file_path, index=False)
        messagebox.showinfo("تم", f"تم حفظ التقرير في {file_path}")
        log_event(f"تم تصدير التقرير: {file_path}")
    except Exception as e:
        messagebox.showerror("خطأ", str(e))
        log_event(f"خطأ تصدير التقرير: {str(e)}")

# جدولة الإرسال
def schedule_sending():
    try:
        start_time_str = entry_schedule.get()
        if not start_time_str:
            start_sending()
            return
        start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
        now = datetime.datetime.now()
        scheduled = now.replace(hour=start_time.hour, minute=start_time.minute, second=0, microsecond=0)
        if scheduled < now:
            scheduled += datetime.timedelta(days=1)
        wait_seconds = (scheduled - now).total_seconds()
        log_event(f"سيبدأ الإرسال في {scheduled.strftime('%H:%M')}")
        messagebox.showinfo("جدولة", f"سيبدأ الإرسال في {scheduled.strftime('%H:%M')}")
        threading.Timer(wait_seconds, start_sending).start()
    except Exception as e:
        messagebox.showerror("خطأ", "صيغة الوقت يجب أن تكون HH:MM")
        log_event(f"خطأ في جدولة الإرسال: {str(e)}")

# إرسال الرسائل
def send_messages():
    global driver
    df = get_filtered_data()
    if df is None or df.empty:
        messagebox.showerror("خطأ", "لا توجد بيانات بعد الفلترة.")
        log_event("لا توجد بيانات بعد الفلترة.")
        return
    try:
        options = Options()
        options.add_experimental_option("detach", True)
        service = Service("msedgedriver.exe")
        driver = webdriver.Edge(service=service, options=options)
        driver.get("https://web.whatsapp.com")
        proceed = messagebox.askokcancel("واتساب", "سجّل دخولك عبر QR ثم اضغط موافق للمتابعة")
        if not proceed:
            driver.quit()
            log_event("تم إلغاء الإرسال قبل البدء.")
            return

        total = len(df)
        progress_bar["maximum"] = total
        sent_count = 0

        for idx, (index, row) in enumerate(df.iterrows(), 1):
            if is_stopped:
                log_event("تم إيقاف الإرسال من قبل المستخدم.")
                break
            while is_paused:
                time.sleep(1)
            phone = str(row["رقم العميل"])
            message = render_template(row["نص الرسالة"], row)
            message_encoded = urllib.parse.quote(message)
            link = f"https://web.whatsapp.com/send?phone={phone}&text={message_encoded}"

            success = False
            for attempt in range(1, max_retries + 1):
                try:
                    driver.get(link)
                    log_event(f"({idx}/{total}) إرسال إلى {phone} - محاولة {attempt}")
                    time.sleep(10)

                    # إرسال المرفقات إن وجدت
                    if attachment_paths:
                        attach_btn = driver.find_element(By.XPATH, '//div[@title="Attach"]')
                        attach_btn.click()
                        time.sleep(2)
                        file_inputs = driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]')
                        for file_input in file_inputs:
                            for path in attachment_paths:
                                file_input.send_keys(path)
                                time.sleep(2)

                    send_btn = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
                    send_btn.click()
                    time.sleep(2)

                    idx_excel = excel_data[(excel_data["رقم العميل"] == row["رقم العميل"])].index[0]
                    excel_data.at[idx_excel, "تقرير الإرسال"] = "تم الإرسال"
                    log_event(f"تم الإرسال إلى {phone}")
                    success = True
                    break
                except Exception as e:
                    log_event(f"خطأ مع {phone}: {str(e)}")
                    if attempt == max_retries:
                        idx_excel = excel_data[(excel_data["رقم العميل"] == row["رقم العميل"])].index[0]
                        excel_data.at[idx_excel, "تقرير الإرسال"] = f"فشل: {str(e)}"
            sent_count += 1
            progress_bar["value"] = sent_count
            label_progress.config(text=f"تم إرسال {sent_count} من {total}")
            root.update_idletasks()
            time.sleep(float(entry_wait.get()))

        excel_data.to_excel("تقرير_الإرسال.xlsx", index=False)
        log_event("تم حفظ تقرير_الإرسال.xlsx")
        messagebox.showinfo("تم", "تم الإرسال، تم حفظ تقرير_الإرسال.xlsx")
    except Exception as e:
        messagebox.showerror("خطأ", str(e))
        log_event(f"خطأ عام: {str(e)}")

# بدء الإرسال
def start_sending():
    global is_paused, is_stopped
    is_paused = False
    is_stopped = False
    progress_bar["value"] = 0
    label_progress.config(text="")
    thread = threading.Thread(target=send_messages)
    thread.start()

def pause_sending():
    global is_paused
    is_paused = not is_paused
    btn_pause.config(text="استئناف" if is_paused else "إيقاف مؤقت")
    log_event("تم إيقاف الإرسال مؤقتًا." if is_paused else "تم استئناف الإرسال.")

def stop_sending():
    global is_stopped
    is_stopped = True
    messagebox.showinfo("تم", "تم إيقاف الإرسال.")
    log_event("تم إيقاف الإرسال كليًا.")

# تقرير الإرسال
def show_report():
    if excel_data is None:
        return
    df = get_filtered_data()
    total = len(df)
    sent = df["تقرير الإرسال"].str.contains("تم الإرسال", na=False).sum()
    failed = df["تقرير الإرسال"].str.contains("فشل", na=False).sum()
    rate = round((sent / total) * 100, 2) if total else 0
    label_total.config(text=f"👥 إجمالي العملاء: {total}")
    label_sent.config(text=f"✅ تم الإرسال: {sent}")
    label_failed.config(text=f"❌ فشل الإرسال: {failed}")
    label_success_rate.config(text=f"📈 نسبة النجاح: {rate}%")
    log_event(f"عرض التقرير: إجمالي={total}, تم الإرسال={sent}, فشل={failed}, نسبة النجاح={rate}%")

# الواجهة الرسومية
root = Tk()
root.title("WhatsApp Sender - فؤاد")
root.geometry("600x700")

btn_load = Button(root, text="📂 تحميل ملف Excel", command=load_excel)
btn_load.pack(pady=5)

btn_attach = Button(root, text="📎 اختر مرفقات لإرسالها", command=choose_attachments)
btn_attach.pack(pady=5)

frame_filters = Frame(root)
frame_filters.pack(pady=5)
Label(frame_filters, text="🔎 فلترة حسب العمود:").grid(row=0, column=0)
filter_column = StringVar()
filter_dropdown = ttk.Combobox(frame_filters, textvariable=filter_column, state="readonly", width=20)
filter_dropdown.grid(row=0, column=1)
filter_dropdown.bind("<<ComboboxSelected>>", update_filter_values)

Label(frame_filters, text="🧪 القيمة:").grid(row=0, column=2)
filter_value = StringVar()
filter_value_dropdown = ttk.Combobox(frame_filters, textvariable=filter_value, state="readonly", width=20)
filter_value_dropdown.grid(row=0, column=3)

# التحكم في وقت الانتظار بين الرسائل
frame_wait = Frame(root)
frame_wait.pack(pady=5)
Label(frame_wait, text="⏱️ وقت الانتظار بين الرسائل (ثواني):").grid(row=0, column=0)
entry_wait = Entry(frame_wait, width=5)
entry_wait.insert(0, "60")
entry_wait.grid(row=0, column=1)

# جدولة الإرسال
frame_schedule = Frame(root)
frame_schedule.pack(pady=5)
Label(frame_schedule, text="🕒 جدولة الإرسال (ساعة:دقيقة - اختياري):").grid(row=0, column=0)
entry_schedule = Entry(frame_schedule, width=7)
entry_schedule.grid(row=0, column=1)
btn_schedule = Button(frame_schedule, text="⏰ جدولة الإرسال", command=schedule_sending)
btn_schedule.grid(row=0, column=2, padx=5)

btn_start = Button(root, text="▶️ بدء الإرسال الآن", command=start_sending)
btn_start.pack(pady=5)

btn_pause = Button(root, text="⏸️ إيقاف مؤقت", command=pause_sending)
btn_pause.pack(pady=5)

btn_stop = Button(root, text="⏹️ إيقاف كلي", command=stop_sending)
btn_stop.pack(pady=5)

btn_report = Button(root, text="📊 عرض التقرير", command=show_report)
btn_report.pack(pady=5)

btn_export = Button(root, text="💾 تصدير التقرير", command=export_report)
btn_export.pack(pady=5)

# شريط التقدم
progress_bar = ttk.Progressbar(root, orient=HORIZONTAL, length=400, mode='determinate')
progress_bar.pack(pady=10)
label_progress = Label(root, text="")
label_progress.pack()

label_total = Label(root, text="👥 إجمالي العملاء: 0")
label_total.pack()
label_sent = Label(root, text="✅ تم الإرسال: 0")
label_sent.pack()
label_failed = Label(root, text="❌ فشل الإرسال: 0")
label_failed.pack()
label_success_rate = Label(root, text="📈 نسبة النجاح: 0%")
label_success_rate.pack()

# سجل الأحداث
Label(root, text="📝 سجل الأحداث:").pack()
log_text = scrolledtext.ScrolledText(root, height=10, state='disabled')
log_text.pack(fill=BOTH, expand=True, padx=10, pady=5)

root.mainloop()