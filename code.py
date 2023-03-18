# -*- coding: utf8 -*
#Jim Christopoulos
#retrieving eMails from gamil from specific user and then converting them to pdf
################################################################################

#email libraries
import imaplib
import email 
from tqdm import tqdm #tracking progress for the for loop
from fpdf import FPDF #for pdf creation
from bs4 import BeautifulSoup #to extract the neccessary elements

# Variavles to declare your email address, password and the gamil imap
user = 'YOUR_EMAIL_ADDRESS@gmail.com' 
password = 'YOUR_SPECIFIC_APP_PASSWORD' 
imap_url = 'imap.gmail.com'
 
# Create a new PDF file
pdf = FPDF()
# Add a page
pdf.add_page()

# Add a DejaVu Unicode font (uses UTF-8), this is done because you might receive characters that are not UTF-8 and thus an error will occur
# Supports more than 200 languages. For a coverage status see:
# http://dejavu.svn.sourceforge.net/viewvc/dejavu/trunk/dejavu-fonts/langcover.txt
# source : https://pyfpdf.readthedocs.io/en/latest/Unicode/index.html
# Also you will need to download the desired font file and save it in the same directory as the .py file 
pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
pdf.set_font('DejaVu', '', 12)

# Create a conncection to your accoount 
con = imaplib.IMAP4_SSL(imap_url)
con.login(user, password)
con.select('Inbox')

# Get the desired data from the desired sender
result, data = con.search(None, 'FROM', 'SENDERS_EMAIL_ADDRESS')

# Iterate through each message in reverse order
# Using tqdm to display a progress bar
for num in tqdm(data[0].split()[::-1]):
    # Fetch the email message
    result, data = con.fetch(num, '(RFC822)')
    email_message = email.message_from_bytes(data[0][1])

    # Extract the desired information from the email message
    sent_to = email_message['To']
    date = email_message['Date']
    subject = email_message['Subject']

    # Extract the message body
    if email_message.is_multipart():
        # For multipart emails, iterate through each part and find the text/plain part
        for part in email_message.get_payload():
            if part.get_content_type() == 'text/plain':
                body = part.get_payload(decode=True).decode('utf-8')
    else:
        # For non-multipart emails, just extract the text
        body = email_message.get_payload(decode=True).decode('utf-8')

    #Uncomment the next line to see the extracted information
    #print('Sent to:', sent_to)
    #print('Date:', date)
    #print('Subject:', subject)
    #print('Body:', body)
    #print('=' * 50) #using it to distinguish where the a new mail strats
    string = "Sent to : " + sent_to +"\n" + "Date : " + date + "\nSubject : " +subject + "\nMessage : \n" + body + "\n" + '='*50
    ##################################
    pdf.multi_cell(0, 5, txt=string)

#save the pdf
pdf.output("emails.pdf") 

# Close the mailbox and logout
con.close()
con.logout()
