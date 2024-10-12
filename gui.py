# gui.py

from pathlib import Path
from tkinter import (
    Tk, Canvas, Entry, Text, Button, PhotoImage, Toplevel, Label,
    StringVar, Listbox, Scrollbar, END, SINGLE, Frame, RIGHT, Y, LEFT,
    filedialog, messagebox, INSERT, SEL_FIRST, SEL_LAST, simpledialog, Radiobutton, OptionMenu
)
from tkinter import ttk
import os
import json
import smtplib
import threading
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email_utils import (
    is_valid_email, save_sender_credentials, load_sender_credentials,
    save_recipients, load_recipients
)

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("assets/frame0")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


def change_email():
    def update_email():
        new_email = email_entry.get()
        new_login = new_email.split('@')[0]
        new_password = password_entry.get()
        if not new_email or not new_login or not new_password:
            error_label.config(text="Пожалуйста, заполните все поля.")
            return
        if is_valid_email(new_email):
            # Обновляем отображаемый email в основном окне
            canvas.itemconfig(email_text_id, text=new_email)
            # Обновляем текущие данные
            global current_email, current_login, current_password
            current_email = new_email
            current_login = new_login
            current_password = new_password
            # Сохраняем данные отправителя
            save_sender_credentials(current_email, current_login, current_password)
            # Очистка сообщения об ошибке
            error_label.config(text="")
            # Закрываем окно изменения email
            email_window.destroy()
        else:
            # Выводим сообщение об ошибке
            error_label.config(text="Некорректный email адрес. Пожалуйста, исправьте.")

    # Создаем новое окно для ввода данных отправителя
    email_window = Toplevel(window)
    email_window.title("Изменить данные отправителя")
    email_window.geometry("400x300")

    # Поля ввода для Email
    Label(email_window, text="Введите Email отправителя:").pack(pady=5)
    email_var = StringVar(value=current_email)
    email_entry = Entry(email_window, width=30, textvariable=email_var)
    email_entry.pack()


    # Поля ввода для Пароля
    Label(email_window, text="Введите Пароль приложения отправителя:").pack(pady=5)
    password_var = StringVar(value=current_password)
    password_entry = Entry(email_window, width=30, textvariable=password_var, show="*")
    password_entry.pack()

    # Метка для отображения сообщений об ошибке
    error_label = Label(email_window, text="", fg="red")
    error_label.pack()

    # Кнопки подтверждения и отмены
    button_frame = Frame(email_window)
    button_frame.pack(pady=10)
    Button(button_frame, text="Сохранить", command=update_email).pack(side=LEFT, padx=5)
    Button(button_frame, text="Отмена", command=email_window.destroy).pack(side=LEFT, padx=5)


def manage_recipients():
    def update_recipients_listbox():
        recipients_listbox.delete(0, END)
        for recipient in recipients:
            recipients_listbox.insert(END, recipient)

    def add_recipient():
        def save_new_recipient():
            new_recipient = recipient_entry.get()
            if is_valid_email(new_recipient):
                recipients.append(new_recipient)
                update_recipients_listbox()
                save_recipients(recipients)
                add_window.destroy()
            else:
                add_error_label.config(text="Некорректный email адрес.")

        add_window = Toplevel(recipients_window)
        add_window.title("Добавить получателя")
        add_window.geometry("350x150")

        Label(add_window, text="Введите Email получателя:").pack(pady=5)
        recipient_entry = Entry(add_window, width=30)
        recipient_entry.pack()
        recipient_entry.focus_set()

        recipient_entry.bind('<KeyPress>', on_key_press)

        add_error_label = Label(add_window, text="", fg="red")
        add_error_label.pack()

        Button(add_window, text="Сохранить", command=save_new_recipient).pack(pady=5)

    def edit_recipient():
        selected_index = recipients_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        old_recipient = recipients[index]

        def save_edited_recipient():
            edited_recipient = recipient_entry.get()
            if is_valid_email(edited_recipient):
                recipients[index] = edited_recipient
                update_recipients_listbox()
                save_recipients(recipients)
                edit_window.destroy()
            else:
                edit_error_label.config(text="Некорректный email адрес.")

        edit_window = Toplevel(recipients_window)
        edit_window.title("Редактировать получателя")
        edit_window.geometry("350x150")

        Label(edit_window, text="Измените Email получателя:").pack(pady=5)
        recipient_entry = Entry(edit_window, width=30)
        recipient_entry.insert(0, old_recipient)
        recipient_entry.pack()
        recipient_entry.focus_set()

        recipient_entry.bind('<KeyPress>', on_key_press)

        edit_error_label = Label(edit_window, text="", fg="red")
        edit_error_label.pack()

        Button(edit_window, text="Сохранить", command=save_edited_recipient).pack(pady=5)

    def delete_recipient():
        selected_index = recipients_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        del recipients[index]
        update_recipients_listbox()
        save_recipients(recipients)

    # Создаем окно управления списком получателей
    recipients_window = Toplevel(window)
    recipients_window.title("Список получателей")
    recipients_window.geometry("400x400")

    # Создаем рамку для размещения Listbox и Scrollbar
    frame = Frame(recipients_window)
    frame.pack(pady=10, fill='both', expand=True)

    # Список получателей с прикрепленным Scrollbar
    scrollbar = Scrollbar(frame)
    scrollbar.pack(side=RIGHT, fill=Y)

    recipients_listbox = Listbox(
        frame, selectmode=SINGLE, width=50, yscrollcommand=scrollbar.set
    )
    recipients_listbox.pack(side='left', fill='both', expand=True)

    scrollbar.config(command=recipients_listbox.yview)

    update_recipients_listbox()

    # Создаем новый фрейм для кнопок
    buttons_frame = Frame(recipients_window)
    buttons_frame.pack(pady=10)

    # Кнопки управления, расположенные в одну строку
    Button(buttons_frame, text="Добавить", command=add_recipient).pack(side=LEFT, padx=5)
    Button(buttons_frame, text="Редактировать", command=edit_recipient).pack(side=LEFT, padx=5)
    Button(buttons_frame, text="Удалить", command=delete_recipient).pack(side=LEFT, padx=5)


def save_text_and_files():
    # Сохраняем текст из entry_1 и тему письма из subject_entry
    text_content = entry_1.get("1.0", END)
    subject_content = subject_entry.get()
    email_data = {
        'text': text_content,
        'subject': subject_content
    }
    with open('email_data.json', 'w', encoding='utf-8') as f:
        json.dump(email_data, f, ensure_ascii=False, indent=4)
    # Сохраняем список файлов
    with open('attached_files.json', 'w', encoding='utf-8') as f:
        json.dump(attached_files, f, ensure_ascii=False, indent=4)
    # Сохраняем настройки расписания
    schedule_data = {
        'schedule_type': schedule_type.get(),
        'start_date': start_date.get(),
        'start_time': start_time.get(),
        'repeat_period': repeat_period.get(),
        'period_unit': period_unit.get()
    }
    with open('schedule_settings.json', 'w', encoding='utf-8') as f:
        json.dump(schedule_data, f, ensure_ascii=False, indent=4)


def load_text_and_files():
    # Загружаем текст и тему письма
    if os.path.exists('email_data.json'):
        with open('email_data.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            entry_1.insert(END, data.get('text', ''))
            subject_entry.insert(0, data.get('subject', ''))
    # Загружаем список файлов
    if os.path.exists('attached_files.json'):
        with open('attached_files.json', 'r', encoding='utf-8') as f:
            files = json.load(f)
            # Проверяем наличие каждого файла
            for file_path in files:
                if os.path.exists(file_path):
                    attached_files.append(file_path)
                    files_listbox.insert(END, os.path.basename(file_path))
                else:
                    messagebox.showwarning("Файл не найден", f"Файл {file_path} не найден и не был загружен.")

    # Загружаем настройки расписания
    load_schedule_settings()


def load_schedule_settings():
    if os.path.exists('schedule_settings.json'):
        with open('schedule_settings.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            schedule_type.set(data.get('schedule_type', 'now'))
            start_date.set(data.get('start_date', datetime.now().strftime('%Y-%m-%d')))
            start_time.set(data.get('start_time', datetime.now().strftime('%H:%M')))
            repeat_period.set(data.get('repeat_period', ''))
            period_unit.set(data.get('period_unit', 'минута/минут'))
    else:
        schedule_type.set('now')
        start_date.set(datetime.now().strftime('%Y-%m-%d'))
        start_time.set(datetime.now().strftime('%H:%M'))
        repeat_period.set('')
        period_unit.set('минута/минут')


def add_file():
    file_path = filedialog.askopenfilename(
        title="Выберите файл",
        filetypes=(("Все файлы", "*.*"),)
    )
    if file_path:
        if file_path not in attached_files:
            attached_files.append(file_path)
            files_listbox.insert(END, os.path.basename(file_path))
        else:
            messagebox.showinfo("Файл уже добавлен", "Этот файл уже добавлен в список.")


def delete_file():
    selected_index = files_listbox.curselection()
    if selected_index:
        index = selected_index[0]
        del attached_files[index]
        files_listbox.delete(index)
    else:
        messagebox.showwarning("Не выбран файл", "Пожалуйста, выберите файл для удаления.")

# Обработчик событий клавиатуры

def send_email():
    # Проверяем, есть ли получатели
    if not recipients:
        messagebox.showwarning("Нет получателей", "Список получателей пуст.")
        return
    # Получаем тему письма
    email_subject = subject_entry.get().strip()
    if not email_subject:
        messagebox.showwarning("Пустая тема", "Тема письма пустая.")
        return

    # Получаем текст письма
    email_text = entry_1.get("1.0", END).strip()
    if not email_text:
        messagebox.showwarning("Пустое письмо", "Текст письма пуст.")
        return

    schedule_type_value = schedule_type.get()

    if schedule_type_value == 'once':
        # Единоразовая отправка в заданное время
        start_datetime_str = f"{start_date.get()} {start_time.get()}"
        try:
            send_time = datetime.strptime(start_datetime_str, '%Y-%m-%d %H:%M')
        except ValueError:
            messagebox.showwarning("Некорректное время",
                                   "Введите корректные дату и время в формате ГГГГ-ММ-ДД и ЧЧ:ММ.")
            return
        delay = (send_time - datetime.now()).total_seconds()
        if delay < 0:
            messagebox.showwarning("Некорректное время", "Время отправки уже прошло.")
            return
        threading.Timer(delay, lambda: actual_send_email(log_to_interface=True)).start()
        messagebox.showinfo("Письмо запланировано", f"Письмо будет отправлено {start_datetime_str}.")
    elif schedule_type_value == 'repeat':
        # Периодическая отправка
        start_datetime_str = f"{start_date.get()} {start_time.get()}"
        try:
            start_time_dt = datetime.strptime(start_datetime_str, '%Y-%m-%d %H:%M')
        except ValueError:
            messagebox.showwarning("Некорректное время",
                                   "Введите корректные дату и время в формате ГГГГ-ММ-ДД и ЧЧ:ММ.")
            return

        repeat_period_value = repeat_period.get()
        period_unit_value = period_unit.get()

        if not repeat_period_value.isdigit():
            messagebox.showwarning("Некорректный период", "Введите числовое значение периода повторений.")
            return
        repeat_period_value = int(repeat_period_value)

        if repeat_period_value <= 0:
            messagebox.showwarning("Некорректный период", "Период повторений должен быть положительным числом.")
            return

        # Вычисляем интервал в секундах  ['минута/минут', 'час/часов', 'день/дни', 'неделя/недели', 'месяц/месяца']
        period_units_in_seconds = {
            'минута/минут': 60,
            'час/часов': 3600,
            'день/дни': 86400,
            'неделя/недели': 604800,
            'месяц/месяца': 2592000  # Приблизительное значение (30 дней)
        }
        interval = repeat_period_value * period_units_in_seconds.get(period_unit_value, 60)

        delay = (start_time_dt - datetime.now()).total_seconds()
        if delay < 0:
            messagebox.showwarning("Некорректное время", "Время первого запуска уже прошло.")
            return

        # Запускаем поток периодической отправки
        global stop_sending
        stop_sending = False

        def schedule_periodic_send():
            time.sleep(delay)
            threading.Thread(target=periodic_send_email, args=(interval,)).start()

        threading.Thread(target=schedule_periodic_send).start()
        messagebox.showinfo("Письмо запланировано",
                            f"Письмо будет отправлено впервые {start_datetime_str} и затем каждые {repeat_period_value} {period_unit_value}.")
        # Обновляем метку о статусе периодической отправки
        update_periodic_status(True)
    else:
        # Немедленная отправка
        threading.Thread(target=actual_send_email, args=(True, True, True)).start()


def actual_send_email(log_to_interface=False, show_message=False, show_progress=False):
    # Показываем индикатор выполнения, если show_progress равно True
    if show_progress:
        progress_window = Toplevel(window)
        progress_window.title("Отправка писем")
        progress_window.geometry("300x100")
        Label(progress_window, text="Пожалуйста, подождите. Письма отправляются...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(pady=10, padx=20, fill='x')
        progress_bar.start()

    failed_recipients = []
    try:
        # Определяем SMTP-сервер исходя из домена отправителя
        domain = current_email.split('@')[1]
        smtp_port = 587
        if 'gmail.com' in domain:
            smtp_server = 'smtp.gmail.com'
        elif 'yahoo.com' in domain:
            smtp_server = 'smtp.mail.yahoo.com'
        elif 'mail.ru' or 'list.ru' or 'bk.ru' in domain:
            smtp_server = 'smtp.mail.ru'
        elif 'outlook.com' in domain or 'hotmail.com' in domain:
            smtp_server = 'smtp-mail.outlook.com'
        elif 'yandex.ru' in domain or 'ya.ru' in domain:
            smtp_server = 'smtp.yandex.ru'
        else:
            # Пользовательский SMTP-сервер
            smtp_server = simpledialog.askstring("SMTP сервер", "Введите адрес SMTP сервера:")
            smtp_port = simpledialog.askinteger("SMTP порт", "Введите номер порта SMTP сервера:", initialvalue=587)
            if not smtp_server or not smtp_port:
                if show_progress:
                    progress_window.destroy()
                return

        # Подключаемся к SMTP-серверу
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        try:
            server.login(current_login, current_password)
        except smtplib.SMTPAuthenticationError as e:
            if show_progress:
                progress_window.destroy()
            log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ошибка авторизации: {e}")
            if show_message:
                messagebox.showerror("Ошибка авторизации", "Проверьте правильность введенных учетных данных.")
            return
        except Exception as e:
            if show_progress:
                progress_window.destroy()
            log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ошибка подключения к SMTP: {e}")
            if show_message:
                messagebox.showerror("Ошибка", f"Не удалось подключиться к серверу SMTP:\n{e}")
            return

        # Отправляем письмо каждому получателю
        for recipient in recipients:
            msg = MIMEMultipart()
            msg['From'] = current_email
            msg['To'] = recipient
            msg['Subject'] = subject_entry.get().strip()
            msg.attach(MIMEText(entry_1.get("1.0", END).strip(), 'plain', 'utf-8'))

            # Прикрепляем файлы
            for file_path in attached_files:
                try:
                    with open(file_path, 'rb') as f:
                        file_data = f.read()
                        file_name = os.path.basename(file_path)
                    attachment = MIMEBase('application', 'octet-stream')
                    attachment.set_payload(file_data)
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
                    msg.attach(attachment)
                except Exception as e:
                    log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ошибка прикрепления файла {file_path}: {e}")
                    continue  # Переходим к следующему файлу

            # Отправляем письмо
            try:
                server.sendmail(current_email, recipient, msg.as_string())
                log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Письмо успешно отправлено {recipient}")
            except Exception as e:
                failed_recipients.append(recipient)
                log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Не удалось отправить письмо {recipient}: {e}")

        server.quit()

        if show_progress:
            progress_window.destroy()

        # Отображаем сообщение, если нужно
        if show_message:
            if failed_recipients:
                messagebox.showwarning("Отправка завершена с ошибками",
                                       f"Не удалось отправить письма следующим адресатам:\n{', '.join(failed_recipients)}")
            else:
                messagebox.showinfo("Отправка завершена", "Все письма успешно отправлены.")

    except Exception as e:
        if show_progress:
            progress_window.destroy()
        log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Общая ошибка при отправке писем: {e}")
        if show_message:
            messagebox.showerror("Ошибка", f"Произошла ошибка при отправке писем:\n{e}")

def periodic_send_email(interval):
    global stop_sending
    while not stop_sending:
        actual_send_email(log_to_interface=True)
        # Ждем заданный интервал или выхода из цикла
        for _ in range(int(interval)):
            if stop_sending:
                break
            time.sleep(1)
    # Обновляем метку о статусе периодической отправки
    update_periodic_status(False)


def open_schedule_settings():
    schedule_window = Toplevel(window)
    schedule_window.title("Настройка периода отправки")
    schedule_window.geometry("400x400")

    # Функция для переключения видимости полей повторения
    def toggle_repeat_fields():
        if schedule_type.get() == 'repeat':
            repeat_label.pack(pady=5)
            repeat_frame.pack()
        else:
            repeat_label.pack_forget()
            repeat_frame.pack_forget()

    # Выбор типа расписания
    Label(schedule_window, text="Выберите тип расписания:").pack(pady=5)
    Radiobutton(schedule_window, text="Отправить сейчас", variable=schedule_type, value='now',
                command=toggle_repeat_fields).pack(anchor='w')
    Radiobutton(schedule_window, text="Единоразовая отправка в заданное время", variable=schedule_type, value='once',
                command=toggle_repeat_fields).pack(anchor='w')
    Radiobutton(schedule_window, text="Периодическая отправка", variable=schedule_type, value='repeat',
                command=toggle_repeat_fields).pack(anchor='w')

    # Поля для ввода даты и времени первого запуска
    Label(schedule_window, text="Дата первого запуска (ГГГГ-ММ-ДД):").pack(pady=5)
    date_entry = Entry(schedule_window, textvariable=start_date, width=15)
    date_entry.pack()

    Label(schedule_window, text="Время первого запуска (ЧЧ:ММ):").pack(pady=5)
    time_entry = Entry(schedule_window, textvariable=start_time, width=15)
    time_entry.pack()

    # Метка для периода повторений
    repeat_label = Label(schedule_window, text="Период повторений:")
    repeat_label.pack(pady=5)

    # Фрейм для размещения периода и единиц измерения на одной строке
    repeat_frame = Frame(schedule_window)
    repeat_entry = Entry(repeat_frame, textvariable=repeat_period, width=10)
    repeat_entry.pack(side=LEFT)

    unit_options = ['минута/минут', 'час/часов', 'день/дни', 'неделя/недели', 'месяц/месяца']  #[minutes', 'hours', 'days', 'weeks', 'months']
    period_unit.set(period_unit.get() if period_unit.get() in unit_options else 'минута/минут')
    unit_option = OptionMenu(repeat_frame, period_unit, *unit_options)
    unit_option.pack(side=LEFT, padx=5)

    # Изначально скрываем поля периода повторений, если тип расписания не 'repeat'
    if schedule_type.get() == 'repeat':
        repeat_label.pack(pady=5)
        repeat_frame.pack()
    else:
        repeat_label.pack_forget()
        repeat_frame.pack_forget()

    # Кнопки "Сохранить" и "Отмена"
    button_frame = Frame(schedule_window)
    button_frame.pack(pady=10)
    Button(button_frame, text="Сохранить", command=lambda: [save_text_and_files(), schedule_window.destroy()]).pack(side=LEFT, padx=5)
    Button(button_frame, text="Отмена", command=schedule_window.destroy).pack(side=LEFT, padx=5)
def stop_periodic_sending():
    global stop_sending
    stop_sending = True
    log_event(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Периодическая отправка остановлена")


def update_periodic_status(is_active):
    if is_active:
        status_label.config(text="Периодическая отправка: ВКЛЮЧЕНА", fg="green")
    else:
        status_label.config(text="Периодическая отправка: ОТКЛЮЧЕНА", fg="red")


def log_event(message):
    log_listbox.insert(END, message)
    # Автоматически скроллим вниз
    log_listbox.yview_moveto(1)


# Функции для работы с буфером обмена
def paste_text(event):
    try:
        clipboard_content = window.clipboard_get()
        widget = event.widget
        if isinstance(widget, (Entry, Text)):
            widget.insert(INSERT, clipboard_content)
    except Exception as e:
        print(f"Ошибка вставки текста: {e}")

def copy_text(event):
    try:
        widget = event.widget
        if isinstance(widget, (Entry, Text)):
            selected_text = widget.selection_get()
            window.clipboard_clear()
            window.clipboard_append(selected_text)
    except Exception as e:
        print(f"Ошибка копирования текста: {e}")

def cut_text(event):
    try:
        widget = event.widget
        if isinstance(widget, (Entry, Text)):
            selected_text = widget.selection_get()
            window.clipboard_clear()
            window.clipboard_append(selected_text)
            widget.delete("sel.first", "sel.last")
    except Exception as e:
        print(f"Ошибка вырезания текста: {e}")

# Обработчик событий клавиатуры
def on_key_press(event):
    # Логируем состояние, keysym и char для отладки
    # print(f"State: {event.state}, Keysym: {event.keysym}, Char: {event.char}, keycode: {event.keycode}, keysym_num: {event.keysym_num} ")
    # Определяем битовую маску для Control
    CONTROL_MASK = 0x8  # 1000
    # Проверяем, нажата ли клавиша Control
    if (event.state & CONTROL_MASK):
        # Получаем символ в нижнем регистре
        keysym_num = event.keysym_num
        if keysym_num == 1084:
            paste_text(event)
            return "break"
        elif keysym_num == 1089:
            copy_text(event)
            return "break"
        elif keysym_num == 1095:
            cut_text(event)
            return "break"

# Инициализируем главное окно
window = Tk()

window.title("Рассылка электронных писем")
window.geometry("880x950")
window.configure(bg="#ffffff")  # Устанавливаем белый фон

canvas = Canvas(
    window,
    bg="#ffffff",
    height=750,
    width=880,
    bd=0,
    highlightthickness=0,
    relief="ridge"
)

canvas.place(x=0, y=0)

canvas.create_rectangle(
    0.0,
    0.0,
    880.0,
    68.0,
    fill="#FFFFFF",
    outline=""
)

canvas.create_text(
    78.0,  # 78.0
    14.0,
    anchor="nw",
    text="Рассылка электронных писем",
    fill="#000000",
    font=("Montserrat Bold", 32 * -1)
)

# Загрузка изображений (убедитесь, что пути к изображениям корректны)
# Если изображения отсутствуют, закомментируйте или удалите эти строки
try:
    image_image_1 = PhotoImage(
        file=relative_to_assets("image_1.png"))
    image_1 = canvas.create_image(
        297.0,
        117.0,
        image=image_image_1
    )

    image_image_2 = PhotoImage(
        file=relative_to_assets("image_2.png"))
    image_2 = canvas.create_image(
        752.0,
        117.0,
        image=image_image_2
    )
except Exception as e:
    print(f"Ошибка загрузки изображений: {e}")

button_2 = Button(
    text="Изменить",
    borderwidth=0,
    highlightthickness=0,
    command=manage_recipients,
    relief="flat",
    fg="#000000",
    font=("Montserrat Bold", 12 * -1),
    bg="#C6D0F4"
)
button_2.place(
    x=777.0,
    y=87.0,
    width=80.0,
    height=30.0
)

canvas.create_text(
    655.0,
    87.0,
    anchor="nw",
    text="Список\nполучателей",
    fill="#000000",
    font=("Montserrat Bold", 20 * -1)
)

# Поле для ввода темы письма
canvas.create_text(
    42.0,
    177.0,
    anchor="nw",
    text="Тема письма",
    fill="#344ED6",
    font=("Montserrat Bold", 12 * -1)
)

subject_entry = Entry(
    window,
    bd=0,
    bg="#EFF1F9",
    fg="#000716",
    highlightthickness=0,
    font=("Montserrat", 14)
)
subject_entry.place(
    x=42.0,
    y=197.0,
    width=514.0,
    height=25.0
)

# Поле для ввода текста письма с скроллбарами
text_frame = Frame(window)
text_frame.place(
    x=38.0,
    y=257.0,
    width=518.0,
    height=338.0
)

entry_1 = Text(
    text_frame,
    bd=0,
    bg="#EFF1F9",
    fg="#000716",
    highlightthickness=0,
    wrap="word"
)
entry_1.pack(side="left", fill="both", expand=True)

# Добавляем скроллбары
text_scrollbar_y = Scrollbar(text_frame, orient="vertical", command=entry_1.yview)
text_scrollbar_y.pack(side="right", fill="y")
entry_1.configure(yscrollcommand=text_scrollbar_y.set)

canvas.create_text(
    42.0,
    237.0,
    anchor="nw",
    text="Текст письма",
    fill="#344ED6",
    font=("Montserrat Bold", 12 * -1)
)

# Привязываем общий обработчик к событию KeyPress
entry_1.bind('<KeyPress>', on_key_press)

# Список прикрепленных файлов с кнопками
files_frame = Frame(window)
files_frame.place(
    x=659.0,
    y=217.0,
    width=187.0,
    height=378.0
)

files_listbox = Listbox(
    files_frame,
    selectmode=SINGLE,
    bg="#EFF1F9",
    fg="#000716",
    highlightthickness=0
)
files_listbox.pack(side="left", fill="both", expand=True)

files_scrollbar = Scrollbar(files_frame)
files_scrollbar.pack(side="right", fill="y")
files_listbox.config(yscrollcommand=files_scrollbar.set)
files_scrollbar.config(command=files_listbox.yview)

attached_files = []  # Список с полными путями к файлам

canvas.create_text(
    659.0,
    167.0,
    anchor="nw",
    text="Прикрепленные файлы",
    fill="#344ED6",
    font=("Montserrat Bold", 12 * -1)
)

# Кнопки "Добавить" и "Удалить" для списка файлов
add_file_button = Button(
    text="Добавить",
    command=add_file,
    bg="#188423",
    fg="#FFFFFF",
    font=("Montserrat Bold", 12 * -1),
    relief="flat"
)
add_file_button.place(
    x=659.0,
    y=605.0,
    width=90.0,
    height=30.0
)

delete_file_button = Button(
    text="Удалить",
    command=delete_file,
    bg="#D63434",
    fg="#FFFFFF",
    font=("Montserrat Bold", 12 * -1),
    relief="flat"
)
delete_file_button.place(
    x=756.0,
    y=605.0,
    width=90.0,
    height=30.0
)

canvas.create_text(
    42.0,
    87.0,
    anchor="nw",
    text="Email отправителя",
    fill="#000000",
    font=("Montserrat Bold", 12 * -1)
)

# Инициализируем текущие данные отправителя
current_email, current_login, current_password = load_sender_credentials()
if not current_email:
    current_email = "Введите Email отправителя"

# Сохраняем идентификатор текста email для дальнейшего обновления
email_text_id = canvas.create_text(
    42.0,
    107.0,
    anchor="nw",
    text=current_email,
    fill="#000000",
    font=("Montserrat Bold", 22 * -1)
)

# Загрузка изображения (проверьте путь)
try:
    image_image_3 = PhotoImage(
        file=relative_to_assets("image_3.png"))
    image_3 = canvas.create_image(
        42.0,
        34.0,
        image=image_image_3
    )
except Exception as e:
    print(f"Ошибка загрузки изображений: {e}")

button_3 = Button(
    text="Изменить",
    borderwidth=0,
    highlightthickness=0,
    command=change_email,
    relief="flat",
    fg="#000000",
    font=("Montserrat Bold", 12 * -1),
    bg="#C6D0F4"
)
button_3.place(
    x=486.0,
    y=87.0,
    width=80.0,
    height=30.0
)

# Метка для отображения статуса периодической отправки
status_label = Label(window, text="Периодическая отправка: ОТКЛЮЧЕНА", fg="red", font=("Montserrat Bold", 12 * -1))
status_label.place(
    x=38.0,
    y=615.0
)

# Кнопка настройки периода отправки
button_schedule = Button(
    text="Настройка периода",
    command=open_schedule_settings,
    bg="#344ED6",
    fg="#FFFFFF",
    font=("Montserrat Bold", 12 * -1),
    relief="flat"
)
button_schedule.place(
    x=38.0,
    y=665.0,
    width=150.0,
    height=30.0
)

# Кнопка остановки периодической отправки
button_stop_sending = Button(
    text="Остановить отправку",
    command=stop_periodic_sending,
    bg="#D63434",
    fg="#FFFFFF",
    font=("Montserrat Bold", 12 * -1),
    relief="flat"
)
button_stop_sending.place(
    x=200.0,
    y=665.0,
    width=150.0,
    height=30.0
)

# Кнопка отправки письма
button_send_email = Button(
    text="Отправить письмо",
    command=send_email,
    bg="#188423",
    fg="#FFFFFF",
    font=("Montserrat Bold", 14 * -1),
    relief="flat"
)
button_send_email.place(
    x=643.0,
    y=646.0,
    width=219.0,
    height=55.0
)

# Переменные для хранения настроек расписания
schedule_type = StringVar(value='now')
start_date = StringVar(value=datetime.now().strftime('%Y-%m-%d'))
start_time = StringVar(value=datetime.now().strftime('%H:%M'))
repeat_period = StringVar(value='1')
period_unit = StringVar(value='минута/минут')

# Флаг для остановки периодической отправки
stop_sending = False

# Загружаем список получателей
recipients = load_recipients()

# Загружаем текст, тему и файлы
load_text_and_files()

# Добавляем лог событий
log_frame = Frame(window)
log_frame.place(
    x=38.0,
    y=725.0,
    width=518.0,
    height=200
)

log_label = Label(log_frame, text="Лог событий:", font=("Montserrat Bold", 12 * -1))
log_label.pack(anchor='w')

log_listbox = Listbox(
    log_frame,
    bg="#EFF1F9",
    fg="#000716",
    highlightthickness=0
)
log_listbox.pack(fill='both', expand=True)

log_scrollbar = Scrollbar(log_frame)
log_scrollbar.pack(side=RIGHT, fill=Y)
log_listbox.config(yscrollcommand=log_scrollbar.set)
log_scrollbar.config(command=log_listbox.yview)

# Обрабатываем событие закрытия окна для сохранения данных
def on_closing():
    save_text_and_files()
    window.destroy()

window.protocol("WM_DELETE_WINDOW", on_closing)

window.resizable(False, False)
window.mainloop()