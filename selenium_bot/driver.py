import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from purchase_bot import Phone
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from seleniumwire import webdriver
from seleniumwire import undetected_chromedriver as uc
from selenium.webdriver.common.proxy import Proxy, ProxyType
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#from bot.middlewares.logging import logger
from notifiers.logging import NotificationHandler
from loguru import logger
from time import sleep
import pandas as pd
import requests
from locator import Locator
from random import uniform
import pyautogui
import random
from fake_useragent import UserAgent
import re
from traceback import print_exc


def delay():
    sleep(uniform(2,3.5))

apiKey = '36AAe8553090dc2669bcf8c51ff52f82'
service = 'sg'
country = 'ru'
#1570084669

def send_photo(chat_id, image_path, image_caption=""):
    data = {"chat_id": chat_id, "caption": image_caption}
    url = f"https://api.telegram.org/bot6501489744:AAEVS3aoQG1JsbkJBWRdTWw7__JuWJvF43w/sendPhoto?chat_id={chat_id}&caption={image_caption}"
    with open(image_path, "rb") as image_file:
        ret = requests.post(url, data=data, files={"photo": image_file})
    return ret.json()

def send_msg(chat_id, image_path, image_caption=""):
    data = {"chat_id": chat_id, "caption": image_caption}
    url = f"https://api.telegram.org/bot6501489744:AAEVS3aoQG1JsbkJBWRdTWw7__JuWJvF43w/sendPhoto?chat_id={chat_id}&caption={image_caption}"
    with open(image_path, "rb") as image_file:
        ret = requests.post(url, data=data, files={"photo": image_file})
    return ret.json()

class OzonParse:
    """Класс парсера"""
    def __init__(self,
                 driver = None,
                 url = None,
                 action = None,
                 article = None,
                 keyword = None,
                 adress = None,
                 excel_index = None,
                 df = None,
                 ):
        self.url = url
        self.driver = driver
        self.action = action
        self.article = article
        self.keyword = keyword
        self.adress = adress
        self.excel_index = excel_index
        self.df = df

    def __accept_cookies(self):
        try:
            button = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="layoutPage"]/div/div/div/div/div/button')))
            self.__click(button)
        except Exception as e:
            logger.error(f"Ошибка при одобрении куки - {e}")

    def get_main_page(self):
        self.driver.get(self.url)
        try:
            if WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//h1[text()='Доступ ограничен']"))):
                self.driver.quit()
        except:
            pass
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.SEARCH_BAR))
        sleep(3)
        pyautogui.moveTo(362,265)
        pyautogui.leftClick()
        self.main()

    def main(self):
        self.__accept_cookies()
        self.__auth()
        self.__add_adress()
        self.driver.switch_to.default_content()
        cards = self.__search(self.keyword)
        self.__random_hover()
        self.__get_random_card(cards)
        self.__add_to_cart()
        self.driver.close()
        delay()
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[0])
        card = self.__search_card(self.article)
        self.__hover(card)
        self.__click(card)
        self.__switch_to_new_window()
        self.__add_to_cart()
        delay()
        self.__go_to_pay()
        
    def __input(self, element, text, delay=0):
        try:
            self.__click(element)
            for character in text:
                self.action.send_keys(character)
                self.action.pause(uniform(0.2+delay, 0.4+delay))
        except Exception as e:
            logger.error(f"Ошибка при вводе текста - {e}")
        self.action.perform()

    def __hover(self, element, offset=0):
        self.driver.execute_script("arguments[0].style.border='3px solid red'", element)
        """Переносит курсор к элементу с повторной проверкой координат перед перемещением."""
        
        # Получаем положение и размеры окна браузера
        window_position = self.driver.get_window_position()
        window_x = window_position['x']
        window_y = window_position['y']
        
        # Получаем смещение прокрутки страницы
        scroll_x = self.driver.execute_script("return window.pageXOffset;")
        scroll_y = self.driver.execute_script("return window.pageYOffset;")

        # Повторно проверяем положение элемента
        element_location = element.location
        element_size = element.size
        element_x = element_location['x'] + element_size['width'] / 2
        element_y = element_location['y'] + element_size['height'] / 2

        # Вычисляем точные координаты элемента на экране
        target_x = window_x + element_x - scroll_x
        target_y = window_y + element_y - scroll_y + offset

        # Плавно перемещаем курсор к элементу
        pyautogui.moveTo(target_x, target_y, duration=random.uniform(1, 2))
        sleep(random.uniform(0.5, 1))



    def __click(self, element):
        try:
            self.action.move_to_element(element)
            self.action.pause(uniform(0.3, 0.5))
            self.action.click()
            self.action.perform()
        except Exception as e:
            logger.error(f"Ошибка при клике - {e}")

    def __search(self, text):
        try:
            search_bar = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(Locator.SEARCH_BAR))
            delay()
            self.__input(search_bar, str(text))
            self.__click(WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(Locator.SEARCH_BUTTON)))
            cards = WebDriverWait(self.driver, 10).until(EC.visibility_of_all_elements_located(Locator.SEARCH_RESULTS))
            return cards
        except Exception as e:
            logger.error(f"Ошибка при поиске - {e}")
    
    def __scroll_to_element(self, element):
        try:
            target_y = element.location['y']
            current_y = 0
            while current_y < target_y:
                current_y += 10
                self.driver.execute_script(f"window.scrollTo(0, {current_y});")
                sleep(0.01)
            sleep(random.uniform(0.5, 1.5))
            self.action.move_to_element(element).perform()
        except Exception as e:
            logger.error(f"Ошибка при скроллинге - {e}")

    def __search_card(self, article: str):
        try:
            while True:
                cards = WebDriverWait(self.driver, 10).until(EC.visibility_of_all_elements_located(Locator.SEARCH_RESULTS))
                for card in cards:
                    link_element = card.find_element(*Locator.HREF)
                    href = link_element.get_attribute("href")
                    article_number = re.search(r'/product/.*-(\d+)/', href).group(1)
                    if str(article_number) == str(article):
                        return card

                # Прокручиваем вниз, если карточка не найдена
                self.driver.execute_script("window.scrollBy(0, 1000);")
                sleep(1)  # Задержка для подгрузки новых элементов
                
                # Проверка: достигнут ли конец страницы
                new_height = self.driver.execute_script("return document.body.scrollHeight;")
                if new_height == self.driver.execute_script("return window.innerHeight + window.pageYOffset;"):
                    break

            return None  # Вернуть None, если карточка так и не найдена
        except Exception as e:
            logger.error(f"Ошибка при поиске карточки по артикулу - {e}")


    def __switch_to_new_window(self):
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])

    def __get_random_card(self, cards):
        try:
            card = random.choice(cards)
            self.__scroll_to_element(card)
            self.__hover(card)
            pyautogui.leftClick()
            pyautogui.leftClick()
            sleep(1)
            self.__switch_to_new_window()
        except Exception as e:
            logger.error(f"Ошибка при получении рандомной карточки - {e}")

    def __random_hover(self):
        """Перемещает курсор в рандомное место в рамках браузера"""
        window_size = self.driver.get_window_size()
        window_width = window_size['width']
        window_height = window_size['height']
        x_offset = random.randint(0, window_width)
        y_offset = random.randint(0, window_height)
        pyautogui.moveTo(x_offset, y_offset, duration=random.uniform(1, 2))
        sleep(random.uniform(0.5, 1))

    def __add_to_cart(self):
        try:
            print(self.driver.current_url)
            name_for_cart = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(Locator.NAME))
            self.name_for_cart = name_for_cart.text.strip()
            try:
                button = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(Locator.ADD_TO_CART1))
            except:
                button = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(Locator.ADD_TO_CART2))
            try:
                pictures = WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located(Locator.PICTURES))[1:]
                delay()
                picture_1 = random.choice(pictures)
                pictures.remove(picture_1)
                self.__hover(picture_1, 100)
                pyautogui.leftClick()
                pyautogui.leftClick()
                try:
                    print(picture_1.parent)
                    picture_2 = random.choice(pictures)
                    pictures.remove(picture_2,100)
                    self.__hover(picture_2)
                    pyautogui.leftClick()
                    pyautogui.leftClick()
                except:
                    print("Нет второй фотки")
                delay()
            except:
                logger.exception("Картинки в карточке товара не найдены")
            self.__hover(button, 100)
            pyautogui.leftClick()
            pyautogui.leftClick()
        except Exception as e:
            logger.error(f"Ошибка при добавлении в корзину - {e}")

    def __go_to_pay(self):
        try:
            cart = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.CART))
            self.__hover(cart, 100)
            self.__click(cart)
            self.__switch_to_new_window()
            checkbox_all = WebDriverWait(self.driver, 10).until(EC.visibility_of_all_elements_located(Locator.CHECKBOX_IN_CART))[0]
            self.__click(checkbox_all)
            checkbox = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, f"//span[contains(text(), '{self.name_for_cart}')]/../../../../../../div/div/label")))
            self.__hover(checkbox, 91)
            self.__click(checkbox)
            sleep(2)
            go_to_pay = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.GO_TO_PAY))
            self.__hover(go_to_pay, 100)
            self.__click(go_to_pay)
            self.__generate_pay()
        except Exception as e:
            logger.error(f"Ошибка при оплате - {e}")

    def __auth(self):
            auth = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.AUTH))
            self.__click(auth)
            sleep(1)
            iframe = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.AUTH_IFRAME))
            self.driver.switch_to.frame(iframe)
            number_field = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.NUMBER_FIELD))
            apiPhone = Phone(apiKey=apiKey, service=service, country=country)
            number = str(apiPhone.get_phone())[1:]
            self.number = number
            self.df.loc[self.excel_index, 'Аккаунт с которого выполнен выкуп'] = str(f"{number}")
            self.df.to_excel('file.xlsx', index=False)
            print(number)
            self.__click(number_field)
            self.__input(number_field, number)
            delay()
            self.__click(WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.BUTTON_AUTH)))
            try:
                try:
                    is_banned = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.CODE_BANNED))
                except:
                    is_banned = False
                if is_banned:
                    self.driver.refresh()
                    self.__auth()
                else:
                    try:
                        code_field = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.CODE_FIELD))
                    except:
                        print('code_field 1 non passed')
                    try:
                        code_field = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.CODE_FIELD_2))
                    except:
                        print('code_field 2 non passed')
                    if self.__check_code_type():
                        try:
                            confirm_eula = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.CONFIRM_EULA))
                            self.__click(confirm_eula)
                        except:
                            print("Нет галочки")
                        sms = apiPhone.get_sms()
                        if sms != None:
                            self.__click(code_field)
                            self.__input(code_field, sms)
                            try:
                                button = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button[type="submit"]')))
                                self.__click(button)
                            except:
                                print("Что-то не так")
                                print_exc()
                            try:
                                name = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="firstName"]')))
                                lastname = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="lastName"]')))
                                button_next = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="patchUserAccountV2"]/button/div[1]')))
                                self.__input(name, "Иван")
                                self.__input(lastname, "Иванов")
                                self.__click(button_next)
                            except:
                                print("Что-то пошло не так при регистрации")
                                print_exc()
                            try:
                                long_time = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="layoutPage"]/div/div/div/div[2]/span')))
                                if long_tim.text == 'Вас давно не было':
                                    self.driver.refresh()
                                    self.__auth()
                            except:
                                print("Успешно")
                        else:
                            self.driver.refresh()
                            self.__auth()
            except Exception as e:
                print(e)
                self.__auth()

    def __add_adress(self):
        self.driver.get("https://ozon.ru")
        adress = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.ADRESS_MAIN))
        self.__click(adress)
        sleep(1)
        adress_choose = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.ADRESS_CHOOSE))
        self.__click(adress_choose)
        sleep(1)
        try:
            script = """
            var element = arguments[0];
            element.focus();
            element.setSelectionRange(0, element.value.length);
            """
            self.driver.switch_to.default_content()
            sleep(2)
            way = str(self.df.at[self.excel_index, 'Способ доставки (Курьером\Самовывоз)'])
            if way == "Курьером":
                courier = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.COURIER))
                self.__click(courier)
                sleep(2)
            try:
                addres_clear = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.ADRESS_CLEAR))
                self.__click(addres_clear)
            except:
                pass
            adress=WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.ADRESS))
            sleep(4)
            self.__input(adress, self.adress)
            self.__hover(adress, 120)
            pyautogui.leftClick()
            pyautogui.leftClick()
            if way == "Курьером":
                adress_confirm = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(Locator.ADRESS_CONFIRM1))
            else:
                adress_confirm = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(Locator.ADRESS_CONFIRM_SAMOVIVOZ))
            sleep(3)
            minus_zoom = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(Locator.MINUS_ZOOM))
            self.__click(minus_zoom)
            sleep(1)
            self.__click(adress_confirm)
            sleep(2)
            if way == "Курьером":
                flat_el = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located(Locator.FLAT))
                entrance_el = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located(Locator.ENTRANCE))
                floor_el = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located(Locator.FLOOR))
                doorphone_el = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located(Locator.DOORPHONE))
                comment_courier_el = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located(Locator.COMMENT_COURIER))
                order_fio_el = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located(Locator.ORDER_FIO))
                flat, entrance, floor, doorphone, comment_courier = str(int(self.df.at[self.excel_index, 'Квартира'])), str(int(self.df.at[self.excel_index, 'Подъезд'])), str(int(self.df.at[self.excel_index, 'Этаж'])), str(int(self.df.at[self.excel_index, 'Домофон'])), str(self.df.at[self.excel_index, 'Комментарий курьеру'])
                self.__input(flat_el,flat)
                self.__input(entrance_el,entrance)
                self.__input(floor_el,floor)
                self.__input(doorphone_el,doorphone)
                self.__input(comment_courier_el,comment_courier)
                self.__click(order_fio_el)
                pyautogui.hotkey('command', 'a')
                pyautogui.press('backspace')
                self.__input(order_fio_el,"Иван Иванов")
                adress_confirm = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(Locator.ADRESS_CONFIRM2))
                self.__click(adress_confirm)
        except:
            print("Что-то пошло не так при добавлении адреса")
            print_exc()

    def __generate_pay(self):
        try:
            sleep(1)
            try:
                iframe = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.AUTH_IFRAME))
                self.driver.switch_to.frame(iframe)
            except:
                pass
            name = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="firstName"]')))
            lastname = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="lastName"]')))
            button_next = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="patchUserAccountV2"]/button/div[1]')))
            self.__input(name, "Иван")
            self.__input(lastname, "Иванов")
            self.__click(button_next)
        except:
            print("Что-то пошло не так при регистрации")
            print_exc()
        try:
            self.driver.switch_to.default_content()
            sleep(2)
            try:
                close = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[4]/button')))
                self.__click(close)
            except:
                print("Нет оплаты после получения")
            try:
                promo = str(self.df.at[self.excel_index, 'Промокод'])
                if promo != "nan":
                    promo_button = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.PROMO_BUTTON))
                    self.__click(promo_button)
                    sleep(1)
                    promo_field = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.PROMO_FIELD))
                    self.__input(promo_field, promo)
                    sleep(1)
                    promo_confirm = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(Locator.PROMO_CONFIRM))
                    self.__click(promo_confirm)
            except:
                print("Нет промокода")
            try:
                sbp = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.FAST_PAY))
            except:
                next_pay = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.NEXT_PAY))
                self.__click(next_pay)
                sleep(1)
                sbp = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.FAST_PAY))
            self.__click(sbp)
            pay = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(Locator.PAY))
            self.__click(pay)
            iframe = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="layoutPage"]/div/div/div/iframe')))
            self.driver.switch_to.frame(iframe)
            qr_code = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.QR_CODE))
            with open('qr_code.png', 'wb') as f:
                f.write(qr_code.screenshot_as_png)
            send_photo(1234746517, "qr_code.png", f"Артикул - {self.article}, Номер телефона - {self.number}")
            send_photo(719383776, "qr_code.png", f"Артикул - {self.article}, Номер телефона - {self.number}")
            self.df.loc[self.excel_index, 'Статус'] = str("Завершено")
            self.df.to_excel('file.xlsx', index=False)

            
            


        except:
            print("Ошибка при генерации QR-кода")
            print_exc()

    def __check_code_type(self):
        try:
            code_type = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.CODE_TYPE)).text
            if "Мы отправили код подтверждения" not in code_type:
                if "цифр" in code_type.split(" "):
                    sleep(25)
                    get_new_code = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.GET_NEW_CODE))
                    self.__click(get_new_code)
                    sleep(60)
                    self.__check_code_type()
                else:
                    return True
            else:
                return True
        except:
            print("Не сброс-звонок")
        try:
            if_email = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(Locator.IF_EMAIL)).text
            print(if_email)
            if "почту" in if_email.split(" "):
                self.__auth()
            else:
                return True
        except:
            print("Не зарегестрирована почта")
        return True
                    
def main():
    url = 'https://ozon.ru'
    window_width = random.randint(1200, 1920)
    window_height = random.randint(850, 1080)
    columns_to_extract = ['Ключевое слово для поиска', 'Артикул', 'Адрес пвз', 'Промокод', 'Статус', 'Аккаунт с которого выполнен выкуп', 'Способ доставки (Курьером\Самовывоз)', "Квартира", "Подъезд","Этаж", "Домофон", "Комментарий курьеру"]
    df = pd.read_excel('file.xlsx', usecols=columns_to_extract)
    df_to_process = df[df['Статус'].isnull()]
    
    wire_options = {
        'proxy': {
            'https': 'https://9oERqg:NgAYyF@147.45.52.51:9424',
        }
    }

    
    for index, row in df_to_process.iterrows():
        retry_count = 3  # Максимальное количество попыток
        attempt = 0
        
        while attempt < retry_count:
            try:
                # Настройка опций Chrome
                options = webdriver.ChromeOptions()
                ua = UserAgent()
                user_agent = ua.random
                options.add_argument("--no-sandbox")
                options.add_argument("--ignore-certificate-errors")
                options.add_argument("--disable-geolocation")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_argument("--disable-extensions")
                options.add_argument("--disable-gpu")
                options.add_argument("--no-first-run")
                options.add_argument("--no-default-browser-check")
                options.add_argument("--disable-notifications")
                options.add_argument("--disable-infobars")
                
                service = ChromeDriverManager(driver_version="131.0.6778.108").install()
                browser_driver = uc.Chrome(seleniumwire_options=wire_options, executable_path=service, options=options)
                
                sleep(0.5)
                browser_driver.set_window_size(window_width, window_height)
                
                driver = OzonParse(
                    url=url,
                    driver=browser_driver,
                    action=ActionChains(browser_driver),
                    article=row['Артикул'],
                    keyword=row['Ключевое слово для поиска'],
                    adress=row["Адрес пвз"],
                    excel_index=index,
                    df=df
                )
                
                driver.get_main_page()  # Попытка выполнить парсинг
                driver.driver.quit()  # Закрыть драйвер после завершения
                
                logger.info(f"Завершен парсинг для URL: {url}")
                break  # Успешное завершение, выходим из цикла попыток
                
            except Exception as error:
                attempt += 1
                print(f"Ошибка при парсинге URL {url}: {error}")
                if attempt < retry_count:
                    print(f"Попытка {attempt} не удалась, пробую снова через 5 секунд...")
                    sleep(5)  # Ожидание перед повторной попыткой
                else:
                    print(f"Не удалось обработать {url} после {retry_count} попыток.")
            finally:
                browser_driver.quit()  # Закрытие браузера в любом случае

if __name__ == "__main__":
    main()