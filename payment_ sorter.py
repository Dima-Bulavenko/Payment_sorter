"""Скрипт создан для поиска емейлов которые содержат квитанции по коммунальным платежам, скачивания этих
квитанций и сортировки из по папкам в зависимости от того чьи они"""

import email
import imaplib
from email.header import decode_header
import os
import fitz
import sys

the_main_path = "C:\Programing\my_program\komunalka\\"
mom_path = "C:\Комуналка\\"
I_path = "C:\Комуналка\\"
payments_name = ["Квартплата",
                 "Вивіз сміття",
                 "доставки газу",
                 "Електропостачання",
                 "Оплата за газ",
                 "Оплата за тепло",
                 "Оплата за водоснабжение",
                 "Обслуговування лічильника теплової енергії",
                 "Управління багатоквартирним будинком"
                 ]

sock = imaplib.IMAP4_SSL("imap.gmail.com", 993)
sock.login('gmail', 'password')

# select the correct mailbox...
sock.select("INBOX")

sock.literal = u"Квитанції за платежами від".encode('utf-8')
typ1, msgsub1 = sock.uid('SEARCH', 'CHARSET', 'UTF-8', 'SUBJECT')
msgsub1 = msgsub1[0].split()

sock.literal = u"Квитанція за платежем від".encode('utf-8')
typ3, msgsub = sock.uid('SEARCH', 'CHARSET', 'UTF-8', 'SUBJECT')
msgsub = msgsub[0].split()

msgsub = msgsub + msgsub1
msgsub.sort()
for i in msgsub:
    res, msg = sock.uid('fetch', i, '(RFC822)')
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            subject, encoding = decode_header(msg['Subject'])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding)
            if "Квитанція за платежем від" in subject or "Квитанції за платежами від" in subject:
                for part in msg.walk():
                    chek_list = []
                    content_disposition = str(part.get("Content-Disposition"))
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        filename, encoding = decode_header(filename)[0]
                        filename = filename.decode(encoding)
                        if ".pdf" in filename:
                            filepath = os.path.join(the_main_path, filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            pdf_document = filepath
                            doc = fitz.open(pdf_document)
                            page_number = doc.pageCount
                            cheak_page_load = doc.loadPage(0)
                            cheak_page_text = cheak_page_load.getText("text")
                            for k in payments_name:
                                if k not in cheak_page_text:
                                    chek_list.append(False)
                                else:
                                    chek_list.append(True)
                            if True in chek_list:
                                if 'Киевская' in cheak_page_text or 'kievskaya' in cheak_page_text:
                                    chosen_path = mom_path
                                elif "Некрасова" in cheak_page_text:
                                    chosen_path = I_path
                                file_name = []
                                for g in range(page_number):
                                    page_load = doc.loadPage(g)
                                    page_text = page_load.getText("text")

                                    for item in payments_name:
                                        if item in page_text:
                                            file_name.append(item)
                                time_for_name = (subject.split())[4]
                                file_name = ("_".join(file_name)).replace('_', ",").replace(' ', '_')
                                file_name = time_for_name + file_name
                                doc.close()
                                os.replace(the_main_path + filename, chosen_path + file_name + ".pdf")
                                if chosen_path == mom_path:
                                    print('Создан файл в папке Мама' + file_name + str(i))
                                elif chosen_path == I_path:
                                    print('Создан файл в папке Я ' + file_name + str(i))
                            elif True not in chek_list:
                                doc.close()
                                os.remove(the_main_path + filename)
                                print("Не комунальный платеж")
