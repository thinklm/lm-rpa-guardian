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
        for file in listdir(self.original_dir):
            if self.padrao_ato.match(file) and self.__created_time_minutes(self.original_dir+file) < self.max_created_time:
                try:
                    print(f"Movendo arquivo {file} -> {self.target_dir}\\ato-inseguro.xlsx")
                    move(self.original_dir+file, f"{self.target_dir}\\ato-inseguro.xlsx")
                except Exception as e:
                    print(e)   
            elif self.padrao_incidente.match(file) and self.__created_time_minutes(self.original_dir+file) < self.max_created_time:
                try:
                    print(f"Movendo arquivo {file} -> {self.target_dir}\\incidente.xlsx")
                    move(self.original_dir+file, f"{self.target_dir}\\incidente.xlsx")
                except Exception as e:
                    print(e)
            elif self.padrao_reconhecimento.match(file) and self.__created_time_minutes(self.original_dir+file) < self.max_created_time:
                try:
                    print(f"Movendo arquivo {file} -> {self.target_dir}\\reconhecimento.xlsx")
                    move(self.original_dir+file, f"{self.target_dir}\\reconhecimento.xlsx")
                except Exception as e:
                    print(e)




if __name__ == "__main__":
    handler = LocalFileHandler()
    handler.move_local_files()