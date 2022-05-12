import smtplib
import time
import os
import imghdr
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

def sendmail(ImgFileName,kime):
    with open(ImgFileName, 'rb') as f:
            img_data = f.read()

    sendto = kime
    user = ''                           # sender mail
    password = ""                       # sender password
    smtpsrv = "smtp.yandex.com.tr"      # server conf
    smtpserver = smtplib.SMTP(smtpsrv, 587)     # server smtp conf


    ### text ###

    ortametin = """ 
    <p>What is Lorem Ipsum? Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.</p>

<p>Why do we use it? It is a long established fact that a reader will be distracted by the readable content of a page when looking at its layout. The point of using Lorem Ipsum is that it has a more-or-less normal distribution of letters, as opposed to using 'Content here, content here', making it look like readable English. Many desktop publishing packages and web page editors now use Lorem Ipsum as their default model text, and a search for 'lorem ipsum' will uncover many web sites still in their infancy. Various versions have evolved over the years, sometimes by accident, sometimes on purpose (injected humour and the like).</p>

<p> Where does it come from? Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur, from a Lorem Ipsum passage, and going through the cites of the word in classical literature, discovered the undoubtable source. Lorem Ipsum comes from sections 1.10.32 and 1.10.33 of &quot;de Finibus Bonorum et Malorum&quot; (The Extremes of Good and Evil) by Cicero, written in 45 BC. This book is a treatise on the theory of ethics, very popular during the Renaissance. The first line of Lorem Ipsum, &quot;Lorem ipsum dolor sit amet..&quot;, comes from a line in section 1.10.32.</p>


    <p><img src="cid:image1"> </p>
    """


    message = MIMEMultipart()
    text = MIMEText(ortametin,'html')
    message['From'] = 'FOO  <foo@foo.com>'         # From
    message['To'] = kime                           # To
    message['Subject'] = 'Subject'                 # Subject
    message.attach(text)
    image = MIMEImage(img_data, name=os.path.basename(ImgFileName))
    message.attach(image)



###     YOU CAN ADD MORE IMAGES AND  PDF FILES  ###


    fp = open('/home/blabla/Documents/bla/foo.jpeg', 'rb') ##  image   path
    msgImage1 = MIMEImage(fp.read(),name=os.path.basename('/home/blabla/Documents/blabla/foo.jpeg'))
    fp.close()
    message.attach(msgImage1)


    fp = open('/home/blabla/Documents/blabla/foo.jpg', 'rb') ##  thumbnail photo path
    thumbnail = MIMEImage(fp.read())
    fp.close()
    thumbnail.add_header('Content-ID','<image1>')
    message.attach(thumbnail)

    fp = open('/home/blabla/Documents/blalba/foo.pdf', 'rb') ##  pdf file path
    part = MIMEBase('application', "octet-stream",name=os.path.basename('/home/blabla/Documents/blabla/foo.pdf'))
    part.set_payload(fp.read())
    encoders.encode_base64(part)
    fp.close()
    message.attach(part)

    smtpserver.ehlo()
    smtpserver.starttls()
    smtpserver.ehlo
    smtpserver.login(user, password)
    (smtpserver.sendmail(user, sendto, message.as_string()))
    smtpserver.close()


email_list = pd.read_excel('/home/blabla/Documents/blabla/foo.xlsx')  ## excel file for email_list
emails = email_list['EMAIL']
for i in range(len(emails)):
    kime = emails[i]
    try:
        sendmail('/home/blabla/Documents/bla/foo.jpeg',kime) # img.src file path
        print(kime + " succesfully sent ")
        time.sleep(13)
    except:
        with open('/home/blabla/Documents/bla/failedmails.txt', 'a') as f: # write failed mails in a excel file
            f.writelines(''.join(kime))
        print(kime + "Failed to send ")
        time.sleep(13)