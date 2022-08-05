import pandas as pd
from datetime import datetime
from os import path, listdir
from shutil import move
import re



class LocalFileHandler:
    def __init__(self, padrao_ato:str=re.compile(r"^unsafe-act.*xlsx$"),
                padrao_incidente:str=re.compile(r"^incident.*xlsx$"),
                padrao_reconhecimento:str=re.compile(r"^acknowledgments.*xlsx$"),
                max_created_time:float=1.0, 
                original_dir:str=r"C:\Users\LMEng\Downloads\\",
                target_dir:str=r"D:\Emanuel\Projetos\EmAndamento\lm-PIs-think\base") -> None:
        self.padrao_ato = padrao_ato
        self.padrao_incidente = padrao_incidente
        self.padrao_reconhecimento = padrao_reconhecimento
        self.max_created_time = max_created_time
        self.original_dir = original_dir
        self.target_dir = target_dir





    def __created_time_minutes(self, file: path) -> float:
        """Check files created time in minutes.

        Args:
            file (path): Path to file to be checked. 

        Returns:
            float: File created time in minutes.
        """
        agora = datetime.now()
        try:
            c_time = datetime.fromtimestamp(path.getctime(file))
            elapsed = agora - c_time

            return elapsed.total_seconds() / 60
        except OSError:
            print(f"Arquivo {file} nao existe ou esta inacessível.")
        except ZeroDivisionError:
            print("Tentativa de divisão por zero!")

        



    def move_local_files (self) -> None:
        """Move recently downloaded files matching string patterns to a
        proper directory with proper filenames.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tMovendo arquivos de {self.original_dir} para {self.target_dir}")

        for file in listdir(self.original_dir):
            if self.padrao_ato.match(file) and self.__created_time_minutes(self.original_dir+file) < self.max_created_time:
                try:
                    print(f"\t\tMovendo arquivo {file} -> {self.target_dir}\\ato-inseguro.xlsx")
                    move(self.original_dir+file, f"{self.target_dir}\\ato-inseguro.xlsx")
                except Exception as e:
                    print(e)

            elif self.padrao_incidente.match(file) and self.__created_time_minutes(self.original_dir+file) < self.max_created_time:
                try:
                    print(f"\t\tMovendo arquivo {file} -> {self.target_dir}\\incidente.xlsx")
                    move(self.original_dir+file, f"{self.target_dir}\\incidente.xlsx")
                except Exception as e:
                    print(e)

            elif self.padrao_reconhecimento.match(file) and self.__created_time_minutes(self.original_dir+file) < self.max_created_time:
                try:
                    print(f"\t\tMovendo arquivo {file} -> {self.target_dir}\\reconhecimento.xlsx")
                    move(self.original_dir+file, f"{self.target_dir}\\reconhecimento.xlsx")
                except Exception as e:
                    print(e)




    def read_excel_new_files (self) -> None:
        """Read new generated excel files and store them in pandas Dataframes.
        """
        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tLendo arquivos gerados")

        try:
            self.df_ato = pd.read_excel(f"{self.target_dir}\\ato-inseguro.xlsx")
            self.df_incidente = pd.read_excel(f"{self.target_dir}\\incidente.xlsx")
            self.df_ack = pd.read_excel(f"{self.target_dir}\\reconhecimento.xlsx")
        except Exception as e:
            print(e)




    def save_to_excel_file (self) -> None:
        """Saving the generated files into worksheets of an unified .xlsx file.
        """
        excel_filepath = "THINK_PIs.xlsx"

        with pd.ExcelWriter(excel_filepath, engine="xlsxwriter") as writer:
            self.df_ato.to_excel(writer, sheet_name="AtoInseguro")
            self.df_incidente.to_excel(writer, sheet_name="Incidente")
            self.df_ack.to_excel(writer, sheet_name="Reconhecimento")

        print(f"[{datetime.now().strftime('%d/%m/%Y')}] -\tArquivos movidos para arquivo único em {excel_filepath}")





if __name__ == "__main__":
    handler = LocalFileHandler()
    # handler.move_local_files()
    handler.read_excel_new_files()
    handler.save_to_excel_file()
    print("OK")
    # handler.unprotect_workbook()