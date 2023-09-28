from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import configparser
import os
from tkinter import *
config = configparser.ConfigParser()
nameConfFile = 'PyVsopenLogin.ini'
#--------------------------------------оболочка---------
window = Tk()
window.title('Добро пожаловать в PyVsopen')
window.geometry('300x500')
textWithCurrentLogin='Сейчас в систему не введены данные ЕСИА'
if nameConfFile in os.listdir():
   config.read(nameConfFile)
   if config['Main']['login'] and config['Main']['password']:
      textWithCurrentLogin='Текущие логин и пароль: '+config['Main']['login']+' '+config['Main']['password']
labelFirst=Label(window, text=textWithCurrentLogin)
loginInput = Entry(window, width=20)
textLogin=Label(window, text="Логин", font=('',16),)
textPass=Label(window, text="Пароль", font=('',16),)
passInput = Entry(window, width=20)
def saveNewLogin():
   if nameConfFile not in os.listdir():
      open(nameConfFile, 'x')
   else:
      open(nameConfFile)
   login = loginInput.get()
   password= passInput.get()
   config['Main']={}#!!!!!!!!!!!!!!!!!!
   config['Main']['login'] = login
   config['Main']['password'] = password
   with open(nameConfFile, 'w') as configFile:
      config.write(configFile)
   loginInput.delete(0, END)
   passInput.delete(0, END)
   textWithCurrentLogin='Текущие логин и пароль: '+config['Main']['login']+' '+config['Main']['password']
   labelFirst.configure(text=textWithCurrentLogin)
btnSaveLogin=Button(text='Сохранить/Обновить', command=saveNewLogin)
btnSaveLogin.grid(column=1, row=3)
labelFirst.grid(row=0, column=0, columnspan=2)
textLogin.grid(column=0, row=1)
loginInput.grid(column=1, row=1)
textPass.grid(column=0, row=2)
passInput.grid(column=1, row=2)
#------------------конце-оболочки----------------
def Main():
   print('Вы всегда можете изменить логин и пароль вручную в файле PyVsopenLogin.ini')
   print('Ожидайте следующих инструкций после полной загрузки страницы...')
   #чтобы открывалось окно firefox нужно либо тут указать путь к файлу geckodriver, либо указать его в PATH
   driver = webdriver.Firefox()
   driver.get("https://belgorod.vsopen.ru/app/login")
   elem = driver.find_element(By.CSS_SELECTOR, '.long_div input')
   elem.click()
   time.sleep(1)
   driver.get(driver.current_url)
   time.sleep(1)
   elem = driver.find_element(By.ID, 'login')
   elem.send_keys(config["Main"]['login'])
   driver.find_element(By.ID, 'password').send_keys(config["Main"]['password'])
   driver.find_element(By.CSS_SELECTOR, '.plain-button.plain-button_wide').click()
   #driver.find_element(By.ID, 'loginByPwdButton').click()
   time.sleep(2)
   driver.get(driver.current_url)
   #Мы вошли!

   while (1):
      print('Откройте журнал, выберите предмет, диапазон дат,  и нажмите "Ввод". Программа заполнит или удалит все уроки от начала до конца выбранного вами диапазона.')
      input()
      #Вошли в таблицу окр мира
      print('Вы хотите заполнить виртуалку, или наоборот - очистить уроки? Введите 1, если первое, 2, если второе.')
      if (input()=='1'):
         print('Введите номер темы урока (используемый системой Pyword/номер темы в excel списке), с которого программа должна начать заполнение данного диапазона')
         current = int(input()) #17
         vsegoDnej = len(driver.find_elements(By.CSS_SELECTOR, '.student__header.week')[1].find_elements(By.CSS_SELECTOR, 'td'))
         for i in range(vsegoDnej): #внимание, объект студент хеадер вик постоянно обновляется, нельзя использовать для него переменную
            if driver.find_elements(By.CSS_SELECTOR, '.student__header.week')[1].find_elements(By.CSS_SELECTOR, '.day_td.js-marks-header.js-create-lesson-button.add-lessons')[i].get_attribute("data-day-with-schedule") == "true":
               if "js-with-lesson" in driver.find_elements(By.CSS_SELECTOR, '.student__lessons.week')[0].find_elements(By.CSS_SELECTOR, '.day_td')[i].get_attribute("class"):
                  driver.find_elements(By.CSS_SELECTOR, '.student__lessons.week')[0].find_elements(By.CSS_SELECTOR, '.day_td')[i].click()
                  oldCreatedLesson=True
               else:
                  oldCreatedLesson=False
                  driver.find_elements(By.CSS_SELECTOR, '.student__header.week')[1].find_elements(By.CSS_SELECTOR, '.day_td.js-marks-header.js-create-lesson-button.add-lessons')[i].click()
               Input = driver.find_element(By.CSS_SELECTOR, '.modal__lessons-input.js-lesson-title.ui-autocomplete-input')
               Input.clear()
               Input.send_keys(str(current))
               time.sleep(0.5)
               if oldCreatedLesson and driver.find_element(By.CSS_SELECTOR, '.js-lesson-homework').get_attribute('value') != '':
                  if current<10:
                     #не ищется по одной цифре
                     text = driver.find_elements(By.CSS_SELECTOR, '.ui-menu-item')[current-1].find_element(By.CSS_SELECTOR,  'a').text
                  else:
                     text = driver.find_element(By.CSS_SELECTOR, '.ui-menu-item').find_element(By.CSS_SELECTOR,  'a').text
                  Input.clear()
                  Input.send_keys(text[:-4])
               else:
                  if current<10:
                     driver.find_elements(By.CSS_SELECTOR, '.ui-menu-item')[current-1].find_element(By.CSS_SELECTOR,  'a').click()
                  else:
                     driver.find_element(By.CSS_SELECTOR,  '.ui-menu-item').find_element(By.CSS_SELECTOR,  'a').click()
               Input.send_keys(Keys.HOME, Keys.DELETE*2)
               if current>=10:
                  Input.send_keys(Keys.HOME, Keys.DELETE)
               if current>=100:
                  Input.send_keys(Keys.HOME, Keys.DELETE)
               driver.find_element(By.CSS_SELECTOR, '.modal__lessons-label').click()
               driver.find_elements(By.CSS_SELECTOR, '.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-draggable.ui-dialog-buttons button')[1].click()
               current+=1
               time.sleep(0.5)
      else:
         vsegoDnej = len(driver.find_elements(By.CSS_SELECTOR, '.student__header.week')[1].find_elements(By.CSS_SELECTOR, 'td'))
         for i in range(vsegoDnej): #внимание, объект студент хеадер вик постоянно обновляется, нельзя использовать для него переменную
            if driver.find_elements(By.CSS_SELECTOR, '.student__header.week')[1].find_elements(By.CSS_SELECTOR, '.day_td.js-marks-header.js-create-lesson-button.add-lessons')[i].get_attribute("data-day-with-schedule") == "true":
               driver.find_element(By.CSS_SELECTOR,  '.student__lessons.week').find_elements(By.CSS_SELECTOR, '.day_td')[i].click()
               driver.find_elements(By.CSS_SELECTOR, '.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-draggable.ui-dialog-buttons button')[2].click()
               driver.switch_to.alert.accept()
               time.sleep(1)
      print('Программа завершена, возможно уроки-контрольные не заполнены')
      print('Вы хотите внести ещё уроки? 1 - да, 2 - нет')
      if (int(input()) == 2):
         driver.close()
         break
btnStart = Button(text='Старт программы!', command=Main)
btnStart.grid(row=4, column=0, columnspan=2, pady=10)
labelInfo = Text(wrap="word", width=30)
infoText = "!Главное окно программы - командная строка. Следите за ней. Программа имеет два режима: удаление всех уроков в выбранной области и автоматическое создание уроков из подгруженного ранее в виртуалку файла excel.\n Вы сможете управлять этим после того, как запустите программу.\nЕсли уроки уже созданы, то программа поменяет только темы в них, не меняя домашнего задания.\nВнимательно выбирайте диапазон дат на сайте и сохраняйте важные данные, прежде чем решите удалить что-то или заполнить новым. Удаление уроков влечёт удаление оценок."
labelInfo.insert('end', infoText)
labelInfo.grid(row=5, column=0, columnspan=2)
window.mainloop()