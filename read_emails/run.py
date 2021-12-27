import email
import time
import imaplib
import dateutil
import traceback
import email.header
import pandas as pd
import requests, json

username = 'email@gmail.com'
password = 'password'


def regexPattern(text,first_text,last_text):
    start = text.find(first_text) + len(first_text)
    end = text.find(last_text)
    substring = text[start:end]
    return substring

def parseHeaders(headers, emDetail):
    for header in headers:
        if header[0] == 'To':
            emDetail['sentTo'] = header[1].replace('<','').replace('>','')
            if ' ' in emDetail['sentTo']:
                emDetail['sentTo'] = emDetail['sentTo'].rsplit(' ',1)[1]
        elif header[0] == 'From':
            emDetail['sentFrom'] = header[1]
        elif header[0] == 'Subject':
            emDetail['subject'] = header[1]
            if 'ISO-8859-1' in emDetail['subject']:
                emDetail['subject'] = str(decode_header(emDetail['subject']))
        elif header[0] == 'Date':
            try:
                date = dateutil.parser.parse(header[1]).strftime('%d-%b-%y')
                emDetail['date'] = date.replace("'",'')
            except:
                print(traceback.format_exc())
                emDetail['date'] = header[1]
    return emDetail


mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(username, password)

startDate = input('Enter start date (e.g 11-17-2020): ')
startDate = dateutil.parser.parse(startDate).strftime('%d-%b-%Y')

mail.select('inbox')
print(f'(SINCE "{startDate}" SUBJECT "SUPPLY order has shipped")' )
result, data = mail.search(None, f'(SINCE "{startDate}" SUBJECT "SUPPLY order has shipped")' )

ids = data[0]
id_list = ids.split()

print(f'Fetched {len(id_list)} emails. Parsing...')
orderList = []
count = 0
for i, each in enumerate(id_list):
    if i % 5 == 0:
        print(f'Parsed {i} of {len(id_list)} emails...')
    try:
        email_id = id_list[i]
        result1, data1 = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(data1[0][1])

        html, emDetail = '', {}
        headers = msg._headers
        emDetail = parseHeaders(headers, emDetail)
        
        for each in msg.get_payload():
            html += str(each.get_payload(i=None, decode=True)).replace('\\r','')

        """
        # emDetail['orderNo'] = regexPattern(html,'ORDER NUMBER:','TRACK YOUR ORDER').strip().replace("\\n","")
        # emDetail['Size'] = regexPattern(html,'SIZE:','QTY:').strip().replace("\\n","")
        # emDetail['size'] = html.split('Size')[1].split('\\n')[0].strip()
        # delivery_addresss = regexPattern(html,'DELIVERY ADDRESS','NEED HELP?').lstrip('\\n').split('\\n',1)
        # emDetail['Name'] = delivery_addresss[0]
        # emDetail['Address'] = delivery_addresss[1].replace("\\n","")
        # tracker_link = regexPattern(html,'TRACK YOUR ORDER','SHIPPED ITEMS').strip().replace("\\n","")
        # tracker_link = tracker_link.replace('<','').replace('>','')
        # tracker_link = requests.get(tracker_link)
        # time.sleep(2)
        # tracker_link = tracker_link.url
        # time.sleep(1)
        # emDetail['tracker_number'] = regexPattern(tracker_link,'&tracknumbers=','&cm_mmc=').strip()
        """

        emDetail['orderNo'] = regexPattern(html.lower(),'order number:','http:').strip().replace("\\n","")
        emDetail['Size'] = regexPattern(html.lower(),'size:','qty:').strip().replace("\\n","")
        emDetail['Product_Name'] = regexPattern(html.lower(),'shipped items','$').strip().replace("\\n","")
        delivery_addresss = regexPattern(html.lower(),'delivery address','via:').lstrip('\\n').split('\\n',1)
        emDetail['Name'] = delivery_addresss[0]
        emDetail['Address'] = delivery_addresss[1].replace("\\n","")
        tracker_link = regexPattern(html.lower(),'http:','track your order').strip().replace("\\n","")
        tracker_link = requests.get('http:'+str(tracker_link))
        time.sleep(3)
        tracker_link = tracker_link.url
        emDetail['tracker_number'] = regexPattern(tracker_link,'&tracknumbers=','&cm_mmc=').strip()
        print('all done..')
    except:
        emDetail['parseStatus'] = 'Parse failure'
        print(traceback.format_exc())
    orderList.append(emDetail)

filename = time.time()
print(f'Parsed all emails. Saving to excel file...')
data_object = pd.DataFrame.from_dict(orderList)
data_object.to_excel('emailOutput_'+str(filename)+'.xlsx')