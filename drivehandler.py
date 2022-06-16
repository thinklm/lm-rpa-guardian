from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

class GDriveHandler:
    def __init__(self, folder_id:str='1FV2QO3cDKVu6LVZt-MsFWubbDWDHSGGG',
                local_base_dir:str = r"D:\Emanuel\Projetos\EmAndamento\lm-PIs-think\base\\") -> None:
        gauth = GoogleAuth(settings_file="settings.yaml")
        gauth.LocalWebserverAuth()
        self.gdrive = GoogleDrive(gauth)
        self.folder_id = folder_id
        self.local_base_dir = local_base_dir




    def upload_files_to_drive (self, file_list:list) -> None:
        try:
            for file in file_list:
                print(f"Uploading file '{file}' to drive...")
                gfile = self.gdrive.CreateFile({
                    'parents': [{'id': self.folder_id}],
                    'title': file
                })
                gfile.SetContentFile(self.local_base_dir+file)
                gfile.Upload()
        except Exception as e:
            print(e)




    def replace_drive_file (self, file: GoogleDrive.CreateFile) -> None:
        try:
            gfile = self.gdrive.CreateFile({
                'parents': [{'id': '1FV2QO3cDKVu6LVZt-MsFWubbDWDHSGGG'}],
                'title': file['title'],
                'id': file['id']
            })
            gfile.SetContentFile(self.local_base_dir+file['title'])
            gfile.Upload()
        except Exception as e:
            print(e)




    def get_drive_files (self) -> None:
        query = f"'{self.folder_id}' in parents and trashed=false"
        return self.gdrive.ListFile({'q': query}).GetList()




if __name__ == "__main__":
    # folder_id ='1FV2QO3cDKVu6LVZt-MsFWubbDWDHSGGG'
    # local_base_dir = r"D:\Emanuel\Projetos\EmAndamento\lm-PIs-think\base\\"
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
                # file_to_replace = gdrive.CreateFile({'id': file['id']})
                gdrive.replace_drive_file(file)
        except Exception as e:
            print(e)