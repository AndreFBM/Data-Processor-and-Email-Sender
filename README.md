# Data-Processor-and-Email-Sender

```markdown
# Automated Data Processor and Email Sender

This script is designed to interact with a database to retrieve order data, connect to an SFTP server to download relevant files, package these files into a zipped folder, and then send an email with the packaged information to specified recipients. This process is intended to automate the workflow for handling promotional material orders, specifically for a system named Promogator.

## Features

- Retrieves data from a specified SQL Server database
- Connects to an SFTP server using provided credentials
- Downloads specified files from the SFTP server
- Zips the downloaded files into a single archive
- Sends the zipped file via email to a list of recipients
- Logs activity and errors into the database

## Prerequisites

To run this script, you need to have the following installed:
- Python 3.x
- Pandas: `pip install pandas`
- Paramiko: `pip install paramiko`
- PyODBC: `pip install pyodbc`
- Python environment with access to an SQL Server
- SMTP server access for sending emails

## Setup

Before running the script, ensure that the following variables are set up according to your environment:

1. Database connection string (`conn`)
2. SFTP credentials (`sftp_credentials`)
3. Sender email and password (`sender`, `emailpass`)
4. Base directory path (`BASE_DIR`)

**WARNING:** Make sure to protect these credentials appropriately.

## Usage

To run the script, execute the main Python file in your command line interface:

```bash
python script_name.py
```

Replace `script_name.py` with the actual file name of the script.

## How It Works

1. The script starts by logging the start of the execution.
2. It retrieves data from the database using `get_data_from_db`.
3. It connects to the SFTP server with `connect_to_sftp`.
4. For each order, it processes and sends data with `process_and_send_data`, which involves:
   - Downloading files
   - Zipping the directory
   - Removing the directory after zipping
   - Sending the email with the zipped file as an attachment
5. Once all orders are processed, the script logs the end of the execution.
6. If any errors occur, they are logged into the database with details.
