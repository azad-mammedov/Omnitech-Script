import random
import hashlib
import urllib.parse
import requests



class SENDSMS():


    unicode = True
    password = 'g#wFcz8a'
    sender = 'Omnitech'
    username = 'omnitech'

    @classmethod
    def sendOTP(cls,phone_number , message):
        phone_number = phone_number
        message = f'OmniTech Test' if not message else message
        cls.sendSMS(phone_number,message)

    @classmethod
    def sendSMS(cls,receiver,text):
        key = cls.create_key(receiver,text)
        params = {
            'login':cls.username,
            'msisdn':receiver,
            'text': text,
            'sender':cls.sender,
            'key':key,
            'unicode':cls.unicode

        }
        url = 'http://apps.lsim.az/quicksms/v1/send?'+urllib.parse.urlencode(params)
        response = requests.get(url)
        response = response.json()
        if response['errorCode']:
            raise Exception(f"SENDSMS ERROR: {response['errorMessage']}")
        return True


    @classmethod
    def create_key(cls,receiver,text):
        hashed_password = hashlib.md5(cls.password.encode('utf-8')).hexdigest()
        key = hashed_password + cls.username + text + receiver + cls.sender
        hashed_key = hashlib.md5(key.encode('utf-8'))
        return hashed_key.hexdigest()

