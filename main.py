from guardianrpa import GuardianDriver
from drivehandler import GDriveHandler
from localfilehandler import LocalFileHandler
from time import sleep
from datetime import datetime



def gather_guardian_data () -> None:
    print("\n\n* Operacoes com Guardian *\n")

    driver = GuardianDriver()
    driver.enter_reports_view()
    driver.select_area()               
    driver.select_date()      

    # Baixa os relatórios por tipo
    for tipo in ["ato", "incidente", "reconhecimento"]:        
        driver.select_report_type(tipo=tipo)   # Seleciona Tipo de Relato
        driver.press_export_button()             # Exporta Arquivo
        sleep(.5)

    sleep(2)
    # Fecha o driver
    driver.quit_driver()




def manage_local_files () -> None:
    print("\n\n* Operacoes com arquivos locais *\n")

    handler = LocalFileHandler()
    handler.move_local_files()




def manage_drive_files () -> None:
    # folder_id ='1FV2QO3cDKVu6LVZt-MsFWubbDWDHSGGG'
    # local_base_dir = r"D:\Emanuel\Projetos\EmAndamento\lm-PIs-think\base\\"
    print("\n\n* Operacoes com arquivos no Drive *\n")

    gdrive = GDriveHandler()

    folder_files = gdrive.get_drive_files()

    if not folder_files:
        upload_file_list = [
            "ato-inseguro.xlsx",
            "incidente.xlsx",
            "reconhecimento.xlsx"
        ]
        gdrive.upload_files_to_drive(file_list=upload_file_list)
    else:
        try:
            for file in folder_files:
                print(f"Replacing file: {file['title']}\t\tid: {file['id']}")
                gdrive.replace_drive_file(file)
        except Exception as e:
            print(e)




if __name__ == "__main__":
    start = datetime.now()

    gather_guardian_data()
    sleep(2)
    manage_local_files()
    manage_drive_files()

    end = datetime.now()
    print(f"\nTempo decorrido: {end - start}\n\n")