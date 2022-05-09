# Send Bulk Email

## Introduction
This small project aims to improve the work efficiency of automatically sending bulk emails along with a spreadsheet.

## Tasks 

 - [x] Every email might have multiple (cc) recipients;
 - [x] Every email has different but similar subject;
 - [x] Every email might have a different message body;
 - [x] Every email should have only one attachment;
 - [x] Every email should include a signature since it's a company-level behaviour.

# Solution

 - **Python** + **smtplib**
**INFO:** The Gmail server is used to test in this small project.

 - Read all the necessary info from an excel file.
**NOTE:** Ensure the excel file that names "Email.xlsx", whose format see below, and the script are in the same directory.

|To|cc|Subject|Body|Attachment|
|--|--|--|--|--|
|  |  |  |  |

- Include an image in the message body as the signature.
**NOTE:** Ensure the PNG file that names "signature.PNG" and the script are in the same directory.

- In the end, this script could be packaged by **PyInstaller**, so that it could run on different OS as required.

# Reference

1. Main reference:
    https://blog.csdn.net/weapon14/article/details/108029602
2. Good example - Python Email Sender: https://github.com/dimaba/sendmail
