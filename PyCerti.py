import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import os
import xlrd


def Mail(toaddr):
    fromaddr = ""  # SENDER'S MAIL
    msg = MIMEMultipart()

    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = ""  # SUBJECT here

    body = ""  # Your mail BODY here

    msg.attach(MIMEText(body, 'plain'))

    filename = 'xyz.jpeg'  # File name with extension to be attached in mail
    # Location of file. Example -> Here sample is a directory in which our file xyz.jpeg exists.
    attachment = open("sample/" + filename,"rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "PASSWORD")  # In place of PASSWORD type sender's password
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()


def CreateCerti(name, college, email):
    img = Image.open("img.jpeg")  # Replace img with name of your Image or Certificate you want to write to
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("font1.ttf", 50)
    # select position to draw text on certificate as per your need
    draw.text((390, 300), name, (0, 0, 0), font=font)
    # Replace font1.ttf with your favorite font file name(keep the font file in the same directory as that of script)
    font = ImageFont.truetype("font1.ttf", 40)
    draw.text((100, 350), college, (0, 0, 0), font=font)

    # This will create a new directory(here named certificates) and then will save all the generated certificates their
    if not os.path.exists('certificates'):
        os.makedirs('certificates')

    img.save('certificates/' + name + '.jpeg', "JPEG", resolution=100.0)
    # Uncomment the below call to mailNow() and this will mail the file for you if you want
    # Mail(email)

    
    
# Replace sheet1.xlsx with your Excel file to be read(keep the file in the same directory as that of script)
loc = ("sheet1.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

for i in range(sheet.nrows):
    # Choose cells as per your need to retrieve data from file. Top left cell is (0,0)
    CreateCerti(sheet.cell_value(i, 1), sheet.cell_value(i, 2),sheet.cell_value(i, 3))
