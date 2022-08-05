## Imports
from guardianrpa import GuardianDriver
from localfilehandler import LocalFileHandler
from time import sleep
from datetime import datetime


## Functions
def gather_guardian_data () -> None:
    """Realiza a raspagem de dados da página do Guardian com bypass de autenticação em cache.
    OS arquivos são baixados em arquivos separados para o diretório de Downloads.

    OBS: É necessário que esteja na rede com acesso à Internet e sem nenhuma página Chrome aberta.
    """
    print(f"[{datetime.now().strftime('%d/%m/%Y')}] -  Operacoes com o Guardian\n")

    driver = GuardianDriver()
    driver.enter_reports_view()
    driver.select_area()               
    driver.select_date()      

    # Baixa os relatórios por tipo
    for tipo in ["ato", "incidente", "reconhecimento"]:        
        driver.select_report_type(tipo=tipo)   
        driver.press_export_button()             
        sleep(2)

    sleep(5)
    driver.quit_driver()




def manage_local_files () -> None:
    """Local files management to update safety PIs database (for Power BI)
    """
    print(f"[{datetime.now().strftime('%d/%m/%Y')}] -  Operacoes com Arquivos Locais\n")

    handler = LocalFileHandler()
    handler.move_local_files()
    handler.read_excel_new_files()
    handler.save_to_excel_file()





def main () -> None:
    """Flow Control Block
    """
    start = datetime.now()    
    print(f"[{datetime.now().strftime('%d/%m/%Y')}] - Iniciando script de atualizacao de base dos PIs THINK:\n\n")

    gather_guardian_data()
    sleep(5)
    manage_local_files()

    end = datetime.now()
    print(f"\nTempo decorrido: {end - start}\n\n")



if __name__ == "__main__":
    main()