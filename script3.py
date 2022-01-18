import os
import json
import sys
import time
from typing import Dict, List
import openpyxl
import requests
import signal
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_DATA = json.loads(open('default_data.json').read())


ITEM_CODES = {
    'TPS320': '98556',
    'NMQ Token':'152152152',
    'TPS320 420': '987654',
    'TPS320. 400':'12345678',
    'Adabtor':'12',
    'Qeydiyyat':'12556858551545554552',
    'Omnitech TPS 575 420':'1234567895285',
    'OMNİTECH TPS575 440':'123456852456',
    'TPS320 410':'123456789',
    'Omnitech TPS 575 (450)':'1234567891001',
    'QEYDİYYATDAN ÇIXMA VƏ DÜŞMƏ':'1234567896357',
    'NKA Təmiri':'1588880002',
    'wollatpos':'666',
    'TPS 320 370':'55'
}





def check_files_exists(file_name : str,*args, **kwargs) ->  None:
    if not os.path.isfile(file_name):
        print('Fayl tapilmadi:', file_name)
        sys.exit()
    


def open_files(files:list , *args, **kwargs) -> 'list[openpyxl.Workbook]':
    opened_files = []
    for  file in files:
        try:
            check_files_exists(file)
            wb = openpyxl.load_workbook(file)
            opened_files.append(wb)
        except Exception as e:
            print('Fayllarin oxusunda error cixdi:',e,'\n','File name:',file,'\n','Developerle elaqe saxlayin')
            sys.exit()
        
    return opened_files

def get_last_row() -> int:
    try:
        last_row = json.loads(open('settings.json').read())['last_row']
        return int(last_row)
    except Exception as e:
        print('Axirinci qaldigimiz setri goturerken error cixdi:',e , '\n' , 'Developerle elaqe saxlayin')
        sys.exit()




def save_workbook(workbook:openpyxl.Workbook,file_name) -> None:
    try:
        workbook.save(file_name)
    except Exception as e:
        print('Workbooku save ederken xeta cixdi:', file_name ,'\n','Error:',e)
        sys.exit()
    







        
ip_address = input('Ip address:')
input_file = input('Oxuyacagimiz faylin adini qeyd edin:') or 'operations_new.xlsx'
output_file = input('Yazacagimiz faylin adini qeyd edin:') or 'results_for_new_operations.xlsx'
files = [input_file , output_file]

row_count = get_last_row()


wb_opened , wb_result  = open_files(files)
ws_opened , ws_result  = wb_opened.active , wb_result.active
settings = json.loads(open('settings.json').read())


def keyboard_exit_handler(*args, **kwargs):
    save_workbook(wb_result , output_file)
    sys.exit('Kod klaviaturadan dayandirildi')

signal.signal(signal.SIGINT , keyboard_exit_handler)
signal.signal(signal.SIGTERM ,keyboard_exit_handler)


def save_setting(settings, status = 'Finished'):
    global row_count
    settings['last_row'] = row_count
    settings['status']  = status
    try:
        open('settings.json','w').write(json.dumps(settings))
    except Exception as e:
        save_workbook(wb_result,output_file)
        sys.exit(f'settingi yazarken xeta cixdi sonuncu setir {row_count}')








def make_data(row , *args, **kwargs) -> dict:
    new_data = DEFAULT_DATA.copy()
    new_data['requestData']['tokenData']['parameters']['data']['items'] = []





    first_item_count = row[9].value or 0
    first_item_price = float(row[10].value or 0)
    first_item_name = row[8].value or ""
    first_item_sum = first_item_price*first_item_count



    if first_item_count and first_item_price and first_item_name:

        first_item_data = {
            'itemCode':ITEM_CODES.get(first_item_name , ''),
            "itemCodeType": 0,
            'itemName':first_item_name,
            'itemPrice':first_item_price,
            'itemQuantity': first_item_count,
            "itemQuantityType": 0,
            'itemSum':first_item_sum,
            "itemVatPercent": 18.0
        }
        new_data['requestData']['tokenData']['parameters']['data']['items'].append(first_item_data)






    second_item_count = row[13].value or 0
    second_item_price = float(row[14].value or 0)
    second_item_name = row[12].value or ""
    second_item_sum = second_item_price * second_item_count
    if second_item_count and second_item_name and second_item_price:
        second_item_data = {
            'itemCode':ITEM_CODES.get(second_item_name , ''),
            "itemCodeType": 0,
            'itemName':second_item_name,
            'itemPrice':second_item_price,
            'itemQuantity':second_item_count,
            "itemQuantityType": 0,
            'itemSum':second_item_sum,
            "itemVatPercent": 18.0
        }
        new_data['requestData']['tokenData']['parameters']['data']['items'].append(second_item_data)



    long_id = row[1].value
    short_id  = long_id[:12]
    refund_document_number = row[0].value

    additional_data = {
        'cashSum':first_item_sum+second_item_sum,

        'parentDocument':long_id,
        'refund_document_number':refund_document_number,
        'refund_short_document_id':short_id,
        'sum':first_item_sum+second_item_sum,
        'vatAmounts':{
            'vatAmount': first_item_sum+second_item_sum,
            'vatPercent':18.0


        }

    }


    # new_data['requestData']['tokenData']['parameters']['data']['items'][0].update(first_item_data)
    # new_data['requestData']['tokenData']['parameters']['data']['items'][1].update(second_item_data)
    new_data['requestData']['tokenData']['parameters']['data'].update(additional_data)


    return json.dumps(new_data)








def send_request(ip_address ,data,port=8989, *args, **kwargs):
    global row_count
    try:
        response = requests.post(f'http://{ip_address}:{port}' , data)

    except Exception as e:
        message = f"Request gonderilken xeta cixdi.{row_count} row-da xeta cixdi iterasiya dayandirildi"
        save_setting(settings,status='Iterasiya qirildi')
        save_workbook(wb_result , output_file)
        sys.exit(message)
   
    return response


def write_result_file(worksheet:Worksheet,row,sended_data , response):
    response_data = dict(response.json())
    request_delivery_status = "Göndərilib"
    if response_data['code'] != 0:
        request_delivery_status =  "Xeta cixdi"
    
    worksheet[row[0].coordinate] = str(sended_data)
    worksheet[row[1].coordinate] = request_delivery_status
    worksheet[row[2].coordinate] = response.text
    save_workbook(wb_result,output_file)
    







def main():
    global row_count

    for row in ws_opened.iter_rows(min_row=2 , max_row=17):
        time.sleep(10)
        data = make_data(row)
        print(data,'\n')
        response = send_request(ip_address , data)
        write_result_file(ws_result , row ,data , response)
        print(response,'saved' , row_count)
        row_count += 1
        time.sleep(5)



if __name__ == '__main__':
    main()