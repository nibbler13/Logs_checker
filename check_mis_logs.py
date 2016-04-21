"""
Check logs file and send email with results
Settings must be in the file check_mis_logs_config.ini at script directory

Author: nibbler13 nn-admin@nnkk.budzdorov.su 31-555
v1.0
2016/04/21
Python version 3.5.1
"""

import configparser
import logging
import smtplib
import email.utils
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from os import path, listdir
from datetime import datetime


def __send_mail(email_server, email_login, email_pass, email_to, message, logfile=''):
    logging.debug("sending email")

    msg_root = MIMEMultipart('mixed')
    msg_root['To'] = email.utils.formataddr(('Recipient', ', '.join(email_to)))
    msg_root['From'] = email.utils.formataddr(('MIS auto-report logs checker', email_login))

    msg_related = MIMEMultipart('related')
    msg_root.attach(msg_related)

    msg_alternative = MIMEMultipart('alternative')
    msg_related.attach(msg_alternative)

    msg_text = MIMEText(message.encode('utf-8'), 'plain', 'utf-8')
    msg_alternative.attach(msg_text)

    msg_html = MIMEText(message.encode('utf-8'), 'html', 'utf-8')
    msg_alternative.attach(msg_html)

    subject = 'All is ok'
    with open(logfile, 'r', encoding='cp1251') as f:
        file_content = f.read()
        if "error" in file_content.lower():
            subject = 'There is some errors in logs'
            msg_attach = MIMEBase('application', 'octet-stream')
            msg_attach.set_payload(file_content.encode())
            encoders.encode_base64(msg_attach)
            msg_attach.add_header('Content-Disposition', 'attachment',
                                  filename=(Header(path.basename(logfile), 'utf-8').encode()))
            msg_root.attach(msg_attach)

    msg_root['Subject'] = subject
    try:
        server = smtplib.SMTP(email_server)
        # server.set_debuglevel(True)
        server.ehlo()

        if server.has_extn('STARTTLS'):
            server.starttls()
            server.ehlo()

        server.login(email_login, email_pass)
        server.send_message(msg_root)
        server.quit()
    except Exception as err:
        logging.error("!!! ERROR: cannot send an email: " + repr(err))


def __check_file(path_to, file_name, reports_name):
    logging.info(file_name)
    message = "<i>" + file_name + "</i><br>"
    tab = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    f_red = '<font color="red">'
    f_size = '<font size="2">'
    f_end = '</font><br>'
    all_ok = True
    with open(path_to + file_name, mode='r', encoding='cp1251') as opened_file:
        while True:
            try:
                current_line = opened_file.readline()
                if not current_line:
                    break
                if 'Execute' in current_line:
                    success = False
                    file_ok = False
                    executed_report = current_line.split(': ')[1][:-1]
                    if executed_report in reports_name:
                        reports_name.remove(executed_report)
                        while True:
                            current_line = opened_file.readline()
                            if not current_line:
                                break
                            if 'Export file' in current_line:
                                file_path = current_line.split(': ')[1][:-1] + ".xls"
                                if path.exists(file_path):
                                    if path.getsize(file_path):
                                        file_ok = True
                                    else:
                                        logging.info("\t!!! ERROR: report '" + executed_report +
                                                     "' exists but have size 0 at: " + file_path)
                                        message += (f_red + tab + executed_report +
                                                    ": файл с отчетом имеет нулевой размер:" + f_end)
                                        message += f_size + tab + '"' + file_path + '"' + f_end
                                else:
                                    logging.info("\t!!! ERROR: report '" + executed_report +
                                                 "' cannot find the exported file: " + file_path)
                                    message += (f_red + tab + executed_report +
                                                ": не удается найти файл с отчетом:" + f_end)
                                    message += f_size + tab + '"' + file_path + '"' + f_end
                            if 'Export completed' in current_line:
                                success = True
                            if '\n' == current_line:
                                if not success:
                                    logging.info("\t!!! ERROR: report '" + executed_report + "' hasn\'t been completed")
                                    message += (f_red + tab + executed_report +
                                                ": отчет не имеет отметки об успешном завершении " + f_end)
                                break
                    if not success or not file_ok:
                        all_ok = False
            except UnicodeError as err:
                logging.error(repr(err))
        if len(reports_name):
            all_ok = False
            logging.info("\t!!! ERROR: cannot find records about the next report: " + repr(reports_name))
            message += f_red + tab + repr(reports_name) + ": отсутствует информация об отчете" + f_end
    if all_ok:
        logging.info("\tall is ok")
        message += tab + "<b>ok</b><br>"
    return message


def check_logs():
    log_file = path.dirname(path.realpath(__file__)) + '\\check_mis_logs.log'
    logging.basicConfig(filename=log_file, level=logging.DEBUG,
                        filemode='w', format='%(asctime)s %(levelname)s: %(message)s',
                        datefmt='%Y.%m.%d %H:%M:%S')

    logging.getLogger().addHandler(logging.StreamHandler())
    logging.debug("started at: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    config_file = path.dirname(path.realpath(__file__)) + '\\check_mis_logs_config.ini'
    config_main_section = 'main'
    config_mail_section = 'mail'

    email_server = '172.16.6.6'
    email_login = 'temp@nnkk.budzdorov.su'
    email_pass = 'Password666'
    email_to = ['s.a.starodymov@7828882.ru']
    message = ''

    def mail():
        __send_mail(email_server, email_login, email_pass, email_to, message, log_file)

    if not path.exists(config_file):
        logging.error("config file '" + config_file + "' doesn\'t exist")
        logging.error("exiting due to error")
        mail()
        exit()

    paths_to_log = list()
    files_to_check = list()
    reports_name = list()

    logging.debug("reading the config file")
    try:
        config = configparser.ConfigParser()
        config.read(config_file, encoding="utf-8")
        paths_to_log = config[config_main_section]['paths_to_log'].splitlines()
        files_to_check = config[config_main_section]['files_to_check'].splitlines()
        reports_name = config[config_main_section]['reports_name'].splitlines()
        email_server = config[config_mail_section]['server_address']
        email_login = config[config_mail_section]['login']
        email_pass = config[config_mail_section]['password']
        email_to = config[config_mail_section]['to'].splitlines()
    except configparser.Error as err:
        logging.error("configparser.Error " + err.message)
    except KeyError as err:
        logging.error("configparser.KeyError " + repr(err))

    if not len(paths_to_log) or not len(files_to_check) or not len(reports_name):
        logging.error("variable 'paths_to_log', 'files_to_check' or 'reports_name' is empty")
        mail()
        exit()

    current_date = datetime.now()
    for p in paths_to_log:
        logging.info("analyzing directory: " + p)
        message += '"' + p + '"<br>'
        if not path.exists(p):
            logging.error("the path '" + p + "' doesn\'t exist")
            continue
        files_in_folders = [f for f in listdir(p) if path.isfile(path.join(p, f))]
        for ftc in files_to_check:
            ftc = str(ftc).split(';')
            found = False
            for f in files_in_folders:
                if ftc[0] in f:
                    f_size = path.getsize(p + f)
                    f_date = path.getatime(p + f)
                    f_dif = (current_date - datetime.fromtimestamp(f_date)).days
                    if f_size and f_dif < int(ftc[1]):
                        found = True
                        message += __check_file(p, f, reports_name.copy())
            if not found:
                logging.info("!!! ERROR: cannot find the file '" + ftc[0] + "'")
                message += '<font color="red">' + ftc[0] + " не удается найти файл</font><br>"
        message += "<br>"

    logging.debug("ended at: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    mail()

if __name__ == '__main__':
    check_logs()
