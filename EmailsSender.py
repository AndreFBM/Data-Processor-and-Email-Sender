import os
import zipfile
import pandas as pd
import paramiko
import pyodbc
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
from datetime import datetime

# Changing Variables
conn = pyodbc.connect('DRIVER={SQL Server};'
                      'SERVER=;'
                      'DATABASE=;'
                      'UID=;'
                      'PWD=')

sftp_credentials = {
    'username': '',
    'password': '',
    'server': '',
    'port': 
}

sender = "@mail.com"
emailpass = ''

BASE_DIR = r'Yourfolder'


def get_data_from_db():
    query = '''select * from wkb_op.dbo.rpa_encomendas_Promogator_View ORDER BY Controlo ASC'''
    df = pd.read_sql(query, conn)
    return df


def connect_to_sftp():
    transport = paramiko.Transport((sftp_credentials['server'], sftp_credentials['port']))
    transport.connect(username=sftp_credentials['username'], password=sftp_credentials['password'])
    sftp = paramiko.SFTPClient.from_transport(transport)
    return sftp, transport


def process_and_send_data(df, sftp):
    for index, row in df.iterrows():
        dir_name = f"{row['Plano']}-{row['OrderlineID']}"
        dir_path = os.path.join(BASE_DIR, dir_name)

        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

        download_files(row, sftp, dir_path)
        zip_directory(dir_name, dir_path)
        shutil.rmtree(dir_path)
        send_email(row, dir_name)
        time.sleep(2)


def download_files(row, sftp, dir_name):
    sftp_path = row['FTP_PATH']
    mockup_file = row['MockupFile']
    artwork_file = row['ArtworkFile']
    etiqueta_abr = row['EtiquetaAbr']

    for file_name in sftp.listdir(sftp_path):
        if file_name == mockup_file or file_name == artwork_file or file_name.startswith(etiqueta_abr):
            sftp.get(os.path.join(sftp_path, file_name), os.path.join(dir_name, file_name))


def zip_directory(dir_name, dir_path):
    zip_file_path = os.path.join(BASE_DIR, f"{dir_name}.zip")
    with zipfile.ZipFile(zip_file_path, 'w') as zip_file:
        for folder, subfolders, files in os.walk(dir_path):
            for file in files:
                zip_file.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder, file), dir_path))


def create_html_table(row):
    table_header = '<table style="border-collapse: collapse; width: 100%;">'
    row_data = [
        row['Plano'], row['OrderlineID'], row['Technique'], row['Quantity'],
        row['Print Color'], row['Client Name'], row['Address'], row['Contact']
    ]
    header_cells = ''.join(
        [f'<th style="border: 1px solid black; padding: 8px; background-color: #4295cb;">{header}</th>' for header in
         row.keys()])
    data_cells = ''.join([f'<td style="border: 1px solid black; padding: 8px;">{data}</td>' for data in row_data])

    return f"{table_header}<tr>{header_cells}</tr><tr>{data_cells}</tr></table>"


def insert_into_POR(row):
    cursor = conn.cursor()
    insert_query = '''
                INSERT INTO DW_360Imprimir.op.Print_OrderRequest 
                (OrderLineID, StockSupplier, PrintSupplier, StockRequestDate, StockOrderQuantity, Store)
                VALUES (?, ?, 'Promogator', GETDATE(), ?, 'NA')
                '''
    cursor.execute(insert_query, row['OrderlineID'], row['StockSupplier'], row['Quantity'])
    conn.commit()


def insert_ActivityLogError(row, e):
    cursor = conn.cursor()
    insert_query = '''
                INSERT INTO wkb_op.dbo.RPA_ActivityLog 
                (RobotType, RobotName, LogMessage, LogDate, LogType, OrderlineID)
                VALUES (2, 'Promogator', ?, GETDATE(), 3, ?)
                '''
    cursor.execute(insert_query, str(e), row['OrderlineID'])
    conn.commit()


def insert_ActivityLog(e, Logtype):
    cursor = conn.cursor()
    insert_query = '''
                INSERT INTO wkb_op.dbo.RPA_ActivityLog 
                (RobotType, RobotName, LogMessage, LogDate, LogType)
                VALUES (2, 'Promogator', ?, GETDATE(), ?)
                '''
    cursor.execute(insert_query, str(e), Logtype)
    conn.commit()


def send_email(row, dir_name):
    zip_file_path = os.path.join(BASE_DIR, f"{dir_name}.zip")
    to_addresses = [addr.strip() for addr in row['EmailTo'].replace(";", ",").split(",")]
    cc_addresses = [addr.strip() for addr in row['EmailCC'].replace(";", ",").split(",")]
    recipients = to_addresses + cc_addresses

    today_date = datetime.today().strftime('%Y-%m-%d')

    subject = f"Bizay Order {today_date} Promogator {row['OrderlineID']}"

    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = ', '.join(to_addresses)
    msg['CC'] = ', '.join(cc_addresses)
    msg['Subject'] = subject

    table = create_html_table(
        row[['Plano', 'OrderlineID', 'Technique', 'Quantity', 'Print Color', 'Client Name', 'Address', 'Contact']])

    greetings = "Good evening,<br><br>"
    new_line = f"In case there's a need for more files, you can find them in the FTP in the following path: {row['FTP_PATH']}<br><br>"
    signature = "<br>Best Regards,<br>Bizay.<br>"

    full_email_content = greetings + new_line + table + signature
    msg.attach(MIMEText(full_email_content, 'html'))

    try:
        with open(zip_file_path, "rb") as attachment:
            part = MIMEBase('application', 'zip')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {dir_name}.zip")
            msg.attach(part)

        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(sender, emailpass)
            server.sendmail(sender, recipients, msg.as_string())

        insert_into_POR(row)
    except Exception as e:
        print(f"Error in send_email: {e}")
        insert_ActivityLogError(row, e)
    finally:
        os.remove(zip_file_path)


def main():
    insert_ActivityLog('Execution started', 1)
    try:
        df = get_data_from_db()
        sftp, transport = connect_to_sftp()
        process_and_send_data(df, sftp)
        sftp.close()
        transport.close()
    except Exception as e:
        print(f"Error in main: {e}")
        insert_ActivityLog(e, 3)
    finally:
        insert_ActivityLog('Execution ended', 2)
        conn.close()


if __name__ == '__main__':
    main()
