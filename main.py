import datetime
import subprocess

while True:
    
    try:
        import configparser
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as ec
        import pandas as pd
        import traceback
        import chromedriver_autoinstaller
        from selenium.webdriver.chrome.service import Service
        import os
        import sys
        import logging
        break

    except ImportError as err_mdl:

        subprocess.check_call([sys.executable, "-m", "pip", "install","--trusted-host", "pypi.org" ,"--trusted-host" ,"files.pythonhosted.org", err_mdl.name])



class AUTO_GOLF_REPLACEMENT:

    def __init__(self):

        self.path = os.getcwd()

        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

        # create console handler
        ch = logging.StreamHandler()

        #create file handler 
        date = str(datetime.datetime.now().strftime('%d-%b-%Y %H_%M_%S %p'))

        fh = logging.FileHandler(f'{self.path}\\debug\\{date}.log',encoding='utf-8')

        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(funcName)s - %(lineno)d - %(levelname)s - %(message)s',datefmt='%d/%b/%Y %I:%M:%S %p')

        # add formatter to ch
        ch.setFormatter(formatter)

        #add formatter to fh
        fh.setFormatter(formatter)

        # add ch to logger
        self.logger.addHandler(ch)

        #add fh to logger
        self.logger.addHandler(fh)


        #config.init file
        self.my_config_parser = configparser.ConfigParser()

        self.my_config_parser.read(f'{self.path}\\config\\config.ini')

        self.config = { 

        'wisitl_usr': self.my_config_parser.get('config','wisitl_usr'),
        'wisitl_pwd': self.my_config_parser.get('config','wisitl_pwd'),
        'yaneew_usr': self.my_config_parser.get('config','yaneew_usr'),
        'yaneew_pwd': self.my_config_parser.get('config','yaneew_pwd'),
        'nathawitp_usr': self.my_config_parser.get('config','nathawitp_usr'),
        'nathawitp_pwd': self.my_config_parser.get('config','nathawitp_pwd'),
        'arissaran_usr': self.my_config_parser.get('config','arissaran_usr'),
        'arissaran_pwd': self.my_config_parser.get('config','arissaran_pwd'),

        }
        
        #init chrome driver
        self.driver_path = chromedriver_autoinstaller.install()
        
        
        #ELEMENT MAPPING
        
        self.value_map = {'Done':68,
                          'Need More Information':23,
                          'Need more info. FAR':261}
        
        
    def read_excel(self,file_path,sheet):
        self.all_cases = { 'wisitl':{},
                            'yaneew':{},
                            'nathawitp':{},
                            'arissaran':{}}

        
        self.df = pd.read_excel(file_path,sheet)

        
        for i in range(self.df.index.stop):
            
            self.all_cases[self.df['Owner'][i]].update({self.df['GOLF ID'][i]: [self.df['Comment to Supplier'][i],self.df['Action'][i]]})
            
            self.logger.info(self.all_cases)
            
        
        
    def login(self,usr,pwd):
        
        self.driver=webdriver.Chrome(self.driver_path)
        self.driver.maximize_window()
        self.driver.get('https://golf.fabrinet.co.th/')


        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(usr)
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]'))).send_keys(pwd)
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="submit"]'))).click()      
        
        
        
    def access_golf(self,owner,golf_id):  
        
        self.driver.get(f'https://golf.fabrinet.co.th/normaluser/WorkFlow.asp?rnt={golf_id}')
        
        frame = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//frame[@name="down"]')))
        
        self.driver.switch_to.frame(frame)
        
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="i0s2c13130"]'))).send_keys(self.all_cases[owner][golf_id][0])
        
        # map action with value
        
        
        value = self.value_map[self.all_cases[owner][golf_id][1]]
        
        
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, f'//input[@name="i0s3c50t14"][@value="{value}"]'))).click()


        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="button"][@name="hello"][@value="Send  Your  Action"]'))).click()

        # handle with alert
        try:
            WebDriverWait(self.driver, 3).until(ec.alert_is_present())

            alert = self.driver.switch_to.alert
            alert.accept()
            self.logger.debug("alert accepted")

        except:
            self.logger.debug("no alert")
        
    
    def main(self):
        
        self.read_excel('golf_replacement.xlsm','Sheet1')
        
        for owner in self.all_cases:
            
            if self.all_cases[owner].__len__() > 0:
            
                self.login(self.config[f'{owner}_usr'],self.config[f'{owner}_pwd'])
                
                for golf in self.all_cases[owner]:
                    
                    self.access_golf(owner,golf)
                    
                self.driver.quit()
            


if __name__ == '__main__':

    try:

        inst = AUTO_GOLF_REPLACEMENT()
        
        inst.main()


    finally:

        inst.logger.critical("Traceback Error: "+traceback.format_exc())



