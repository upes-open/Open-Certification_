from flask import Flask, render_template, request
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('main.html')

@app.route('/create')
def create():
    return render_template('certificate-1.html')


@app.route('/generate', methods=['GET', 'POST'])
def generate():

    def certificate_name(draw,color,name):
        name=name.upper()
        x=550
        y=520
        padding_x=35
        padding_y=5
        font = ImageFont.truetype('Sanchez-Regular.ttf', size=75)
        text_width, text_height = draw.textsize(name, font=font)
        draw.rectangle((x - padding_x, y - padding_y, x + text_width + padding_x, y + text_height + padding_y), fill='white')
        draw.text((x, y), name, fill=color, font=font)
    
    def certificate_organisation(draw,color,organisation):
        x=560
        y=705
        padding_x=70
        padding_y=10
        organisation=organisation.upper()
        text=f'by  {organisation}. We appreciate'
        font2 = ImageFont.truetype('Sanchez-Regular.ttf', size=30)
        text_width, text_height = draw.textsize(text, font=font2)
        draw.rectangle((x - padding_x, y - padding_y, x + text_width + padding_x, y + text_height + padding_y), fill='white')
        draw.text((x, y), text, fill=color, font=font2)

    def certificate_signature_name(draw,color,signature_name):
        x=670
        y=960
        padding_x=10
        padding_y=10
        signature_name=signature_name.upper()
        font2 = ImageFont.truetype('Sanchez-Regular.ttf', size=25)
        text_width, text_height = draw.textsize(signature_name, font=font2)
        draw.rectangle((x - padding_x, y - padding_y, x + text_width + padding_x, y + text_height + padding_y), fill='white')
        draw.text((x, y), signature_name, fill=color, font=font2) 

    def certificate_designation(draw,color,designation):
        x=690
        y=1010
        padding_x=10
        padding_y=10
        designation=designation.upper()
        font2 = ImageFont.truetype('Sanchez-Regular.ttf', size=20)
        text_width, text_height = draw.textsize(designation, font=font2)
        draw.rectangle((x - padding_x, y - padding_y, x + text_width + padding_x, y + text_height + padding_y), fill='white')
        draw.text((x, y), designation, fill=color, font=font2) 
    
    def certificate_event_name(draw,color,event_name):
        x=500
        y=656
        padding_x=110
        padding_y=10
        event_name=event_name.upper()
        font2 = ImageFont.truetype('Sanchez-Regular.ttf', size=30)
        text=f'For actively participating in {event_name} organised'
        text_width, text_height = draw.textsize(text, font=font2)
        draw.rectangle((x - padding_x, y - padding_y, x + text_width + padding_x, y + text_height + padding_y), fill='white')
        draw.text((x, y), text, fill=color, font=font2)

    def add_logo(draw, logo_path):
        x=700
        y=30
        logo_width=150
        logo_height=150
        padding_x=95
        padding_y=20
        logo = Image.open(logo_path)
        logo = logo.resize((logo_width, logo_height))
        draw.rectangle((x - padding_x, y - padding_y, x + logo_width + padding_x, y + logo_height + padding_y), fill='white')
        certificate.paste(logo, (x, y), logo)
    
    def add_signature(draw, signature_path):
        x=750
        y=800
        signature_width=150
        signature_height=150
        padding_x=95
        padding_y=10
        signature = Image.open(signature_path)
        signature = signature.resize((signature_width, signature_height))
        draw.rectangle((x - padding_x, y - padding_y, x + signature_width + padding_x, y + signature_height + padding_y), fill='white')
        certificate.paste(signature, (x, y), signature)

    fromEmail = "your_email@gmail.com"  # Update with your email
    pwd = "your_password"  # Update with your password

    participants_file = request.files.get('participants_file')
    logo_path=request.form.get('org_logo')
    signature_path=request.form.get('signature')
    organisation=request.form.get('organisation')
    signature_name=request.form.get('signature_name')
    designation=request.form.get('designation')
    event_name=request.form.get('event_name')
    if participants_file and participants_file.filename.endswith(('.xlsx', '.xls')):
        participants_data = pd.read_excel(participants_file)

        for i in participants_data.index:
            certificate = Image.open('../static/images/certificate-1.png')
            draw = ImageDraw.Draw(certificate)
            color = 'rgb(0,0,0)'

            name = participants_data['Name'][i]

            certificate_name(draw,color,name)
            certificate_organisation(draw,color,organisation)
            certificate_signature_name(draw,color,signature_name)
            certificate_designation(draw,color,designation)
            certificate_event_name(draw,color,event_name)
            add_logo(draw, logo_path)
            add_signature(draw,signature_path)
            
            certificateName = "Certificate of "+ name + ".pdf"
            certificate.save(certificateName)
            toEmail = participants_data['Email'][i]
            msg = MIMEMultipart()
            msg['From'] = fromEmail
            msg['To'] = toEmail
            msg['Subject'] = "Certificate Hackathon"
            
            body = '''Thank you ''' + name +''' participating in Open Meet
            
            Regards From
            OPEN'''
            msg.attach(MIMEText(body, 'plain'))
            attachmentName = name + ".pdf"
            with open(certificateName, "rb") as f:
                attach = MIMEApplication(f.read(),_subtype="pdf")
            attach.add_header('Content-Disposition','attachment',filename=attachmentName)
            msg.attach(attach)
            s = smtplib.SMTP('smtp.gmail.com', 587)   
            s.starttls()
            s.login(fromEmail, pwd)
            text = msg.as_string()
            s.sendmail(fromEmail, toEmail, text)

            s.quit()
    
    return "Certificate generation process initiated."

if __name__ == '__main__':
    app.run(debug=True)