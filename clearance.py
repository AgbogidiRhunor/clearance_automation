import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    SessionNotCreatedException, StaleElementReferenceException, 
    NoSuchWindowException, NoSuchElementException, TimeoutException
)
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
import pandas as pd
import datetime
from urllib3.exceptions import ReadTimeoutError
import os
import json
import math
from multiprocessing import Process, freeze_support

CONFIG_PATH = './resource/config.json'
DRIVER_PATH = './resource/geckodriver.exe'
RESOURCE_FOLDERS = ['mat_numbers', 'last_enteries']
MAX_RETRY_COUNT = 5
service = Service(executable_path=DRIVER_PATH)

# COLOR FORMATS
HEADER = '\033[95m'
OKBLUE = '\033[94m'
OKCYAN = '\033[96m'
OKGREEN = '\033[92m'
WARNING = '\033[93m'
FAIL = '\033[91m'
ENDC = '\033[0m'
BOLD = '\033[1m'
UNDERLINE = '\033[4m'

class Bot:

    def __init__(self):

        self.config_path = CONFIG_PATH
        self.load_config()
        self.create_folders()
        self.retry_count = 0
        self.active = False

    def load_config(self):
        with open(self.config_path, 'r') as f:
            self.config = json.load(f)

    def create_folders(self):
        
        for folder in RESOURCE_FOLDERS:
            os.makedirs(f'resource/{folder}', exist_ok=True)

    def cleanup_old_session(self):
        
        if not os.listdir('./resource/mat_numbers'):
            return
        
        print(f"🤖{WARNING} [LOG]{ENDC} -> {FAIL}You are about to delete the old session. Kindly confirm below: {ENDC}\n\n")
        
        while True:
            confirm_deletion = input("Confirm session deletion (y/n): ").lower()
            
            if confirm_deletion == 'y':
                for folder in RESOURCE_FOLDERS:
                    for file in os.listdir(f'./resource/{folder}'):
                        os.remove(f'./resource/{folder}/{file}')

                print(f"🤖{OKGREEN} [LOG]{ENDC} -> {WARNING}Cleanup successful {ENDC}\n\n")
                break

            elif confirm_deletion == 'n':
                print(f"🤖{OKGREEN} [LOG]{ENDC} -> {WARNING}Cleanup aborted {ENDC}\n\n")
                break

            else:
                print(f"🤖{FAIL} [LOG]{ENDC} -> {WARNING}Invalid option selected {ENDC}\n\n")

    def configure_new_session(self):
        print(f"🤖{OKGREEN} [LOG]{ENDC} -> {WARNING}Provide the following details below: {ENDC}\n\n")
        
        # validation
        retry = True

        while retry:
            date_month = input("Enter month(eg 03): ")
            date_day = input("Enter day(eg 20): ")
            running_mode = input("Enter mode (chapel or church): ")
            number_of_processes = input("Enter number of processes (Max is 8, Min is 2): ")

            if number_of_processes:
                if not number_of_processes.isnumeric():
                    print(f"🤖{FAIL} [LOG] -> {ENDC}{WARNING}Number of processes must be a numeric value {ENDC}\n\n")
                    retry = True
                    continue

                number_of_processes = int(number_of_processes)
                if number_of_processes < 2 or number_of_processes > 8:
                    print(f"🤖{FAIL} [LOG] -> {ENDC}{WARNING}Maximum number of processes is 8 and minimum is 2{ENDC}\n\n")
                    retry = True
                    continue

            if not (date_month.isnumeric() and date_day.isnumeric()):
                print(f"🤖{FAIL} [LOG] -> {ENDC}{WARNING}Month, day, and number of processes must be numeric values {ENDC}\n\n")
                retry = True
                continue

            if running_mode not in ['chapel', 'church']:
                print(f"🤖{FAIL} [LOG] -> {ENDC}{WARNING}Running mode can only{ENDC} {OKBLUE}chapel{ENDC} or {OKBLUE}church{ENDC}\n\n")
                retry = True
                continue

            retry = False
        
        # read current content of config file
        with open(self.config_path, 'r') as f:
            configuration = json.load(f)

        # update values
        configuration['execution_config']['day'] = date_day
        configuration['execution_config']['month'] = date_month
        configuration['execution_config']['running_mode'] = running_mode

        if number_of_processes:
            configuration['execution_config']['number_of_processes'] = number_of_processes

        # save new changes to config file 
        with open(self.config_path, 'w') as f:
            json.dump(configuration, f)

        self.config = configuration

        print(f"🤖{OKGREEN} [LOG] -> {ENDC}{OKBLUE}Configuration saved successfully{ENDC}")

    def login(self):

        try:
            # open login page
            self.driver.get(self.login_url)

            email_field = self.driver.find_element(By.NAME, "txtEmail")
            email_field.send_keys(self.config['credentials']['email'])
            print(f"🤖{WARNING} [LOG] {ENDC}-> {OKBLUE}Process {self.process_id}{ENDC}{OKGREEN} Entered email successfully.{ENDC}")

            password_field = self.driver.find_element(By.NAME, "txtPassword")
            password_field.send_keys(self.config['credentials']['password'])
            print(f"🤖{WARNING} [LOG] {ENDC}-> {OKBLUE}Process {self.process_id}{ENDC}{OKGREEN} Entered password successfully.{ENDC}")

            password_field.send_keys(Keys.RETURN)
            time.sleep(10)

        except ReadTimeoutError:
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Lost connection to internet.{ENDC}")
            self.handle_retry() 

        except NoSuchElementException:
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Lost connection to internet.{ENDC}")
            self.handle_retry() 
        
        except Exception as e:
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Unknown error (login) for process {self.process_id}: \n\n{e}\n\n{ENDC} {OKBLUE}Rebooting process {self.process_id}...{ENDC}")
            self.handle_retry()
    
          
    def split_and_save_mat_numbers_from_xlsx_file_to_seperate_process_text_files(self):

        xlsx_file_present = False # flag 
        while not xlsx_file_present:

            folder_content = os.listdir('.')
            for file in folder_content:
                if file.endswith('.xlsx'):
                    df = pd.read_excel(file)
                    xlsx_file_present = True

            if not xlsx_file_present:
                print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}No xlsx file found. Kindly paste this file in the base folder{ENDC}\n")
                input("Press enter when you have added file")
 
        # Remove blank and duplicate values
        if 'Log' not in df.columns:
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}'log' column not found in Excel file.{ENDC}")
            return

        df = df.dropna(subset=['Log'])  # remove empty rows
        df['Log'] = df['Log'].astype(str).str.strip()  # clean spaces
        df = df.drop_duplicates(subset=['Log'])  # remove duplicate numbers
        df = df.reset_index(drop=True)

        total_rows = len(df)
        number_of_processes = int(self.config['execution_config']['number_of_processes'])
        chunk_size = math.ceil(total_rows / number_of_processes)

        print(f"Total rows: {total_rows}, Processes: {number_of_processes}, Chunk size: {chunk_size}")
        # time.sleep(1000)
        for i in range(number_of_processes):
            start = i * chunk_size
            end = start + chunk_size
            chunk = df.iloc[start:end]

            file_path = f'resource/mat_numbers/process_{i+1}.txt'
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            with open(file_path, 'w', encoding='utf-8') as f:
                for value in chunk['Log']:
                    f.write(f"{value}\n")

        print(f"✅ Split {total_rows} rows into {number_of_processes} files in 'resource/mat_numbers/' folder.")
            
    def save_last_entry(self, entry):

        with open(f'resource/last_enteries/process_{self.process_id}.txt', 'w') as f:
            f.write(entry)

        # print(f"🤖{WARNING} [LOG] {ENDC}-> {OKGREEN}Saved {entry} as last entry.{ENDC}")

    def load_matric_numbers(self):

        print(f"🤖{WARNING} [LOG] {ENDC}-> {OKGREEN}Loading mat numbers into memory.{ENDC}")
        
        if 'mat_numbers' in os.listdir('./resource'):
            with open(f'resource/mat_numbers/process_{self.process_id}.txt', 'r', encoding='utf-8') as f:
                mat_numbers = f.readlines()
                mat_numbers = [mat_number.strip('\n').replace(" ", "") for mat_number in mat_numbers]

        else:
            self.split_and_save_mat_numbers_from_xlsx_file_to_seperate_process_text_files()
            self.load_matric_numbers()

        if f'process_{self.process_id}.txt' in os.listdir('./resource/last_enteries'):
            with open(f'resource/last_enteries/process_{self.process_id}.txt', 'r') as f:
                entry = f.read()


            mat_numbers = mat_numbers[(mat_numbers.index(entry)):]

        if len(mat_numbers) == 1:
            print(f"🤖{OKGREEN} [LOG] {ENDC}-> {OKBLUE}Process {self.process_id}{ENDC}{OKGREEN} completed input (matric load).{ENDC}")
            self.driver.quit()
            exit()
        else:
            print(f"🤖{WARNING} [LOG] {ENDC}-> {OKGREEN}Process {self.process_id} loaded {len(mat_numbers)} mat numbers into memory.{ENDC}")
            return mat_numbers

    def add_entry(self):

        try:

            self.driver.get(self.chapel_clearance_url)

            print(f"🤖{WARNING} [LOG] {ENDC}-> {OKGREEN}Setting appropriate date.{ENDC}")

            # set the date 
            self.driver.execute_script("""
                const date = document.getElementById('ContentPlaceHolder1_txtDate');
                date.value = arguments[0];
            """, self.date)

            self.active = True
            matric_numbers = self.load_matric_numbers()
            if matric_numbers:
                for log_value in matric_numbers:
                    try:
                    
                        print(f"🤖{WARNING} [LOG] {ENDC}-> {OKBLUE}Process {self.process_id}{ENDC}{OKGREEN} inputing {log_value}.{ENDC}")

                        # Wait for the text input element to be visible
                        search_input = WebDriverWait(self.driver, 500).until(
                            EC.visibility_of_element_located((By.ID, 'ContentPlaceHolder1_txtSearch'))
                        )

                        # Set the value of the text input to the log_value
                        search_input.clear()  # Clear any existing text in the field
                        time.sleep(1)
                        search_input.send_keys(log_value)

                        # Wait for the search button to be clickable
                        search_btn = WebDriverWait(self.driver, 100).until(
                            EC.element_to_be_clickable((By.ID, 'ContentPlaceHolder1_btnSearch'))
                        )

                        # Click the search button
                        search_btn.click()

                        time.sleep(3)

                        # wait for checkbox to appear and click it
                        checkbox = WebDriverWait(self.driver, 100).until(
                            EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_gdvDetails_chkSelect_0'))
                        )
                        checkbox.click()

                        time.sleep(1)

                        # click submit button
                        submit_btn = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_btnMove'))
                        )
                        submit_btn.click()

                        self.save_last_entry(log_value)
                    
                    except StaleElementReferenceException:
                        continue

                    except NoSuchWindowException:
                        continue

                    except (ReadTimeoutError, TimeoutException):
                        print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Process {self.process_id}.{ENDC} found no record for {log_value}{OKBLUE}{ENDC}")
                        continue

                print(f"🤖{OKGREEN} [LOG] {ENDC}-> {OKBLUE}Process {self.process_id}{ENDC}{OKGREEN} completed input (add entry).{ENDC}")
                self.driver.quit()
                exit()

        except (TimeoutError, ReadTimeoutError):
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Process {self.process_id} timed out.{ENDC} {OKBLUE}Rebooting...{ENDC}")
            self.handle_retry()

        except Exception as e:
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Unknown error (add entry) for process {self.process_id}: \n\n{e}\n\n{ENDC} {OKBLUE}Rebooting process {self.process_id}...{ENDC}")
            self.handle_retry()

    def handle_retry(self):
        print(f"🤖{WARNING} [LOG] {ENDC}-> {WARNING}Rebooting process {self.process_id}.{ENDC}{OKBLUE} Make sure you are connected to stable internet{ENDC}")
        
        if self.retry_count != MAX_RETRY_COUNT:
            self.retry_count += 1
            self.run(retry=True, process_id=self.process_id)
        else:
            print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}Max retry reached for Process {self.process_id}.{ENDC}{OKBLUE} Connect to stable internet and try again{ENDC}")
            self.driver.quit()

    def run(self, process_id, retry=False):

        if retry and self.active:
            self.driver.quit()

        self.process_id = process_id
        # self.load_options()

        execution_config = self.config['execution_config']
        # print(execution_config)
        self.date = f"{datetime.datetime.now().year}-{execution_config['month']}-{execution_config['day']}"
        # 2025-03-30

        if execution_config['running_mode'] == 'chapel':
            self.chapel_clearance_url = self.config['links']['okha_chapel_clerance_url']
        else:
            self.chapel_clearance_url = self.config['links']['faith_arena_chapel_clerance_url']
            
        self.login_url = self.config['links']['login_url']

        firefox_options = webdriver.FirefoxOptions()
        try:
            self.driver = webdriver.Firefox(service=service, options=firefox_options)
        except SessionNotCreatedException as e:
            print(f"🤖{FAIL} [LOG] {ENDC}-> Geckodriver does not match your current firefox version. \nVisit the following url to download the latest version\n\nhttps://geckodriver.com/download/{WARNING}{ENDC}")
            self.handle_retry()

        self.login()
        self.add_entry()

def start_bot(process_id):
    bot = Bot()
    bot.run(process_id)

if __name__ == '__main__':

    freeze_support() 

    print(r"""
   ______     ______   ______   ______     __   __     _____     ______     __   __     ______     ______    
/\  __ \   /\__  _\ /\__  _\ /\  ___\   /\ "-.\ \   /\  __-.  /\  __ \   /\ "-.\ \   /\  ___\   /\  ___\   
\ \  __ \  \/_/\ \/ \/_/\ \/ \ \  __\   \ \ \-.  \  \ \ \/\ \ \ \  __ \  \ \ \-.  \  \ \ \____  \ \  __\   
 \ \_\ \_\    \ \_\    \ \_\  \ \_____\  \ \_\\"\_\  \ \____-  \ \_\ \_\  \ \_\\"\_\  \ \_____\  \ \_____\ 
  \/_/\/_/     \/_/     \/_/   \/_____/   \/_/ \/_/   \/____/   \/_/\/_/   \/_/ \/_/   \/_____/   \/_____/ 
                                                                                                           
                                                                                                    
          """)

    while True:
        split_mat_numbers = os.listdir('./resource/mat_numbers')
        if not split_mat_numbers:
            print(f"🤖{WARNING} [LOG] -> Operations: {ENDC}\n\n{OKGREEN}1. New attendance session{ENDC}\n\n")
        else:
            print(f"🤖{WARNING} [LOG] -> Operations: {ENDC}\n\n{OKGREEN}1. New attendance session\n2. Resume current attendance{ENDC}\n\n")
        
        operation = input("Select an operation: ")
        print('\n')

        if operation == "1":
            bot = Bot()
            bot.cleanup_old_session()
            bot.configure_new_session()
            bot.split_and_save_mat_numbers_from_xlsx_file_to_seperate_process_text_files()
            break

        elif operation == "2":   
            if not split_mat_numbers:
                print(f"🤖{FAIL} [LOG] {ENDC}-> {WARNING}No mat numbers found. Kindly execute operation 1 {ENDC}\n")
            else:
                break

        else:
            print(f"🤖{FAIL} [LOG]{ENDC} -> {WARNING}Invalid operation {ENDC}\n\n")
                
    # Load config once
    with open(CONFIG_PATH, 'r') as f:
        config = json.load(f)

    number_of_processes = config['execution_config']['number_of_processes']

    # Start multiple processes
    processes = []
    for i in range(number_of_processes):
        p = Process(target=start_bot, args=(i+1,))
        processes.append(p)
        p.start()

    for p in processes:
        p.join()
