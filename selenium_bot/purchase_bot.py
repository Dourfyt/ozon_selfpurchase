from requests import get
import time
import logging

apiKey = '36AAe8553090dc2669bcf8c51ff52f82'
service = 'sg'
country = 'ru'

logging.basicConfig(
    level=logging.INFO,  # Уровень логирования
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',  # Формат сообщения
    handlers=[
        logging.FileHandler("app.log"),  # Запись в файл
        logging.StreamHandler()          # Вывод в консоль
    ]
)

# Создаем логгер
logger = logging.getLogger(__name__)

class Phone():
    def __init__(self,apiKey,service,country):
        self.apiKey = apiKey
        self.service = service
        self.country = country
        self.number = None
        self.id = None

    def get_phone(self):
        try:
            resp = get(f"https://api.sms-activate.ae/stubs/handler_api.php?api_key={self.apiKey}&action=getNumberV2&service={self.service}&country={self.country}&maxPrice=20&userId=11472129").json()
            self.number = resp["phoneNumber"]
            self.id = resp["activationId"]       
            return self.number
        except:
            logger.info("Ошибка при получении номера")

    def get_sms(self):
        for attempt in range(15):
            try:
                resp = get(f"https://api.sms-activate.ae/stubs/handler_api.php?api_key={self.apiKey}&action=getStatusV2&id={self.id}").json()
                sms = resp["sms"]["code"]
                return sms
            except Exception as e:
                print(e)
                time.sleep(1)
        return None

if __name__ == "__main__":
    api = Phone(apiKey=apiKey, service=service, country=country)
    api.get_phone()
    api.get_sms()