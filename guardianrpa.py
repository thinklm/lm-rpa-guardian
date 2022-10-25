from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from datetime import datetime
import os
from dotenv import load_dotenv
load_dotenv()


class GuardianDriver:
    def __init__(self) -> None:    # OK   
        self.driver = self.__init_driver()
        self._base_url = os.environ.get("BASE_URL")
        




    def __init_driver(self) -> webdriver.chrome.webdriver.WebDriver:
        """Setup and init Chrome Webdriver

        Returns:
            webdriver.chrome.webdriver.WebDriver: Chrome Webdriver instance to handle automations.
        """
        # print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tIniciando Webdriver")

        options = webdriver.ChromeOptions()
        options.add_argument(r"--user-data-dir=C:\Users\LMEng\AppData\Local\Google\Chrome\User Data")
        options.add_argument(r"--profile-directory=Profile 1")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)




    def __bypass_auth (self) -> None:
        """Bypass Microsoft SSO chached on --user-dir if token expired (daily). 
        It requires the user to confirm again on Authenticator App monthly.
        """
        # print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tRealizando o bypass no SSO da Microsoft")

        account_xpath = '//*[@id="tilesHolder"]/div[1]/div/div[1]'
        try:
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, account_xpath))
            )
            element = self.driver.find_element(By.XPATH, account_xpath)

            action = ActionChains(self.driver)
            action.move_to_element(element).perform()
            action.click().perform()
        except Exception as e:
            print(f"Erro ao bypassar 2FA: {e}")




    def __guardian_login(self) -> None:
        """Logs into Guardian Website via SSO.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tRealizando o login no Guardian")

        try:
            token_notification_xpath = "/nz-notification/div/div/div/div/div[2]"
            bt_ABInBev_xpath = "/html/body/guardian-root/ng-component/nz-spin/div/div/div[2]/guardian-login-selector/div/div/div[1]"
            WebDriverWait(self.driver, 5).until(
                EC.invisibility_of_element((By.XPATH, token_notification_xpath))
            )
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, bt_ABInBev_xpath))
            )
            element = self.driver.find_element(By.XPATH, bt_ABInBev_xpath)

            action = ActionChains(self.driver)
            action.move_to_element(element).perform()
            action.click().perform()

        except Exception as e:
            print(f"Erro no Login: {e}")
            return None

        finally:
            self.driver.implicitly_wait(1)
            sleep(1)
            if self.driver.current_url == f"{self._base_url}home":
                print("Login realizado com sucesso!")
            elif self.driver.current_url == f"{self._base_url}login":
                print("Ainda na página de Login!")
            elif "https://login.microsoftonline.com/" in self.driver.current_url:
                print("Login Automático Falhou!")
                self.__bypass_auth()




    def enter_reports_view (self) -> None:
        """Access Reports view page from the website.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tEntrando na view de relatorios do Guardian")
        
        try:
            self.driver.get(f"{self._base_url}report")  
            self.driver.implicitly_wait(0.5)

        except Exception as e:
            print(f"Exceção ao entrar na view Relatórios: {e}")
            print(f"URL Final: {self.driver.current_url}")

        finally:
            if self.driver.current_url == f"{self._base_url}login":
                self.__guardian_login()
                self.enter_reports_view()
            



    def select_date (self) -> None:
        """Select date range for the reports.
        It was set up to gather the current year data.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tSelecionando intervalo de datas dos relatorios")

        primeiro_dia_ano = datetime.today().replace(day=1, month=1)
        data_inicial_xpath = "/html/body/guardian-root/ng-component/div/div/guardian-reports/div/form/div/div/div/div[1]/div[1]/div[1]/nz-form-item/nz-form-control/div/div/nz-range-picker/div[1]/input"
        data_final_xpath = "/html/body/guardian-root/ng-component/div/div/guardian-reports/div/form/div/div/div/div[1]/div[1]/div[1]/nz-form-item/nz-form-control/div/div/nz-range-picker/div[3]/input"
        
        try:
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, data_inicial_xpath))
            )
            element = self.driver.find_element(By.XPATH, data_inicial_xpath)
            action = ActionChains(self.driver)

            action.move_to_element(element).click()
            action.send_keys(primeiro_dia_ano.strftime("%d/%m/%Y"))
            action.perform()

        except Exception as e:
            print(f"Erro na data inicial: {e}")
            print(f"URL Final: {self.driver.current_url}")

        try:
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, data_final_xpath))
            )
            element = self.driver.find_element(By.XPATH, data_final_xpath)
            action = ActionChains(self.driver)

            action.move_to_element(element).click()
            action.send_keys(datetime.today().strftime("%d/%m/%Y"))
            action.perform()

        except Exception as e:
            print(f"Erro na data final: {e}")
            print(f"URL Final: {self.driver.current_url}")




    def select_area (self) -> None:
        """Select Reports Area field and sets it to "Todas as Áreas"
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tSelecionando \"Todas as Areas\"")

        # Abrir o menu dropdown
        area_xpath = "/html/body/guardian-root/ng-component/div/div/guardian-reports/div/form/div/div/div/div[1]/div[1]/div[2]/div/nz-select/nz-select-top-control/nz-select-search"
        try:
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, area_xpath))
            )
            element = self.driver.find_element(By.XPATH, area_xpath)

            action = ActionChains(self.driver)
            action.move_to_element(element).click()
            action.send_keys("Todas as Áreas")
            action.perform()

        except Exception as e:
            print(f"Erro ao selecionar uma área: {e}")
            print(f"URL Final: {self.driver.current_url}")


        # Selecionar "Todas as Áreas"
        all_areas_xpath = '//*[contains(text(), "Todas as Áreas")]'
        try:
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, all_areas_xpath))
            )
            element = self.driver.find_element(By.XPATH, all_areas_xpath)

            action = ActionChains(self.driver)
            # action.move_to_element(element).click()
            action.move_to_element(element)
            action.click()
            action.perform()

        except Exception as e:
            print(f"Erro ao clicar em todas as áreas: {e}")
            print(f"URL Final: {self.driver.current_url}")




    def select_report_type (self, tipo:str="ato") -> None:
        """Choose the report type from the options.

        Args:
            tipo (str, optional): Report type ("ato", "incidente, "reconhecimento"). Defaults to "ato".
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tSelecionando tipo de relato: {tipo}")

        relato_map = {"ato": 1, "incidente": 3, "reconhecimento": 4}
        relato_xpath = f"/html/body/guardian-root/ng-component/div/div/guardian-reports/div/form/div/div/div/div[1]/div[2]/guardian-radio-button-group/form/div/div/nz-form-control/div/div/nz-radio-group/label[{relato_map[tipo]}]/span[1]"
            
        try:
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, relato_xpath))
            )
            element = self.driver.find_element(By.XPATH, relato_xpath)

            action = ActionChains(self.driver)
            action.move_to_element(element).click().perform()
        except Exception as e:
            print(f"Erro ao selecionar tipo de relato: {e}")




    def press_export_button (self) -> None:
        """Click on export button to generate csv/xlsx file.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tBaixando Arquivo")
        
        bt_export_xpath = "/html/body/guardian-root/ng-component/div/div/guardian-reports/div/form/div/div/div/div[2]/div[2]/button"
        try:
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, bt_export_xpath))
            )
            self.driver.find_element(By.XPATH, bt_export_xpath).click()

        except Exception as e:
            print(f"Erro ao exportar CSV: {e}")



    
    # def __finish_downloads (self) -> None:
    #     if not self.driver.current_url.startswith ("chrome://downloads"):
    #         self.driver.get("chrome://downloads/")
    #     return self.driver.execute_script("""
    #         var items = document.querySelector('downloads-manager')
    #             .shadowRoot.getElementById('downloadsList').items;
    #         if (items.every(e => e.state === "COMPLETE"))
    #             return items.map(e => e.fileUrl || e.file_url);
    #         """)




    # def wait_downloads_finish (self) -> None:
    #     paths = WebDriverWait(self.driver, 120, 1).until(self.__finish_downloads)
    #     print(paths)




    def quit_driver (self) -> None:
        """Close and quit Chrome Webdriver.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] - Saindo do Driver do Chrome.")
        self.driver.close()
        self.driver.quit()





if __name__ == "__main__":
    # Cria o webdriver do chrome
    driver = GuardianDriver()

    # Acessa a view de relatórios e realizar login SSO, se necessário
    driver.enter_reports_view()

    driver.select_area()

    # Configurações de Data dos Relatórios
    driver.select_date()      

    # Baixa os relatórios por tipo
    for tipo in ["ato", "incidente", "reconhecimento"]:        
        driver.select_report_type(tipo=tipo)
        driver.press_export_button()
        sleep(.5)
    sleep(.5)

    # TESTE - Espera os downloads acabarem
    # driver.wait_downloads_finish()

    # Fecha o driver
    driver.quit_driver()
