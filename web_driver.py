# author: Williams Bobadilla
# creado: 4 marzo 2022
# editado por: Williams Bobadilla
# editado: 7 marzo 2022
# descripcion: Manejador de descargas para el driver de Microsoft Edge


import winreg
import requests
import zipfile
import io
import traceback
from selenium import webdriver
from time import sleep
import xmltodict, json
import logging
import datetime
import logging.config
from O365 import Account, FileSystemTokenBackend, MSGraphProtocol
import socket
import shutil

class EdgeDriverLocal:
    
    PATH_OF_PROPERTIES = 'properties.json'
    today_datetime = datetime.datetime.now()
    today = today_datetime.strftime('%d%m%Y%H')
    LOG_FORMAT = '%(levelname)s %(asctime)s - %(message)s'
    logging.config.dictConfig({
    'version': 1,
    'disable_existing_loggers': True,
    })
    logging.basicConfig(filename=f'{today}_log.log',
                        level = logging.DEBUG,
                        format = LOG_FORMAT,
                        filemode = 'w')
    logger = logging.getLogger(__name__)


    def __init__(self, os_type = "win64"):
        """
        Class to fetch the same driver version as your edge browser and 
        save it locally, copy to bots folders, and if the main version doesn't work, it
        try another versions for the driver. 
        
            Params: 
                os_type (str): current machine OS["win64","linux64","win32"]
            Returns: 
                EdgeDriverLocal (obj): object to control download, test, etc. of the driver
        """
        self.version = self._get_edge_version()
        self.os_type = os_type
        self.url = f"https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/{self.version}/edgedriver_{self.os_type}.zip"
        
        #properties 
        properties = self.import_data()
        self.DRIVER_NAME = properties.get("driver_name")
        self.BASE_URL = properties.get("base_url") 
        self.EMAILS_TO_REPORT = properties.get("emails_to_report")
        self.CURRENT_PATH = properties.get("current_path")
        self.TOKEN_ID = properties.get("token_id")
        self.HOST = properties.get("host")
        self.TENANT_ID = properties.get("tenant_id")
        self.CLIENT_ID = properties.get("client_id")


        self.print_and_log("*"*40,"info")
        self.print_and_log("Driver Manager Started","info")
        self.print_and_log("*"*40,"info")
        self.print_and_log(f"Edge version: {self.version} OStype: {self.os_type}")


    def _get_edge_version(self):
        """
        Get the version of the microfost Edge installed in the current machine.

            Params: 
                None
            Returns: 
                edge_version (str): version of edge browser 
        """
        key_path = r"Software\Microsoft\Edge\BLBeacon"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        edge_version = winreg.QueryValueEx(key, "version")[0]
        return edge_version


    def download(self):
        """
        Method to download the zip and put it into the folder to use.
            Params: 
                None
            Returns:
                r.ok (bool): True or False if the response was succesful 
        """
        try:
            self.print_and_log(f"Download started with url: {self.url}","info")
            r = requests.get(self.url, stream=True)
            if r.status_code == 200:
                z = zipfile.ZipFile(io.BytesIO(r.content))
                z.extractall()
                self.print_and_log(f"Download finished with status: {r.ok}","info")
                return r.ok
            self.print_and_log(f"Not found : {self.url}","warning")
            return False
        except Exception as e:
            self.print_and_log(f"Download failed with with url: {self.url}","error")
            self.print_and_log(str(e))
            self.print_and_log(traceback.format_exc())
            return False
    

    def test(self):
        """
        Method to test the downloaded driver.
            Params:
                None
            Retrurns: 
                status (bool): status of the test, if it worked, true, if not, false.
        """
        try:
            self.print_and_log(f"Testing started","info")
            driver = webdriver.Edge(executable_path=self.DRIVER_NAME) 
            sleep(1) 
            driver.close()
            self.print_and_log(f"Testing finished succesfully","info")
            return True
        except Exception as e: 
            self.print_and_log(f"Test failed, have to try another driver version...","warning")
            self.print_and_log(str(e))
            self.print_and_log(traceback.format_exc())
            return False

        
    def try_another_version(self):
        """
        Method to find another version that match the browser version.

            Params: 
                none 
            Returns: 
                satus (bool): status of the download, if it worked, true, if not, false.
        """
        try:
            self.print_and_log(f"Downloading another version started","info")
            res = requests.get(self.BASE_URL)
            obj = xmltodict.parse(res.text)
            results = json.dumps(obj)
            results = json.loads(results)
            urls = results.get("EnumerationResults").get("Blobs").get("Blob")
            list_urls = [ u for u in urls if u["Name"].split(".")[0] == self.version.split(".")[0] and self.os_type in u["Name"]]
            # let's try another driver, download it and test
            for obj in list_urls:
                self.print_and_log(obj)
                self.url = obj.get("Url") # modify the url to download it 
                download_status = self.download()
                if download_status: 
                    test_status = self.test()
                    if test_status: # if it works, break the loop
                        self.print_and_log(f"Download finished and working of another version with url: {self.url}","info")
                        version = obj.get("Name")
                        self.print_and_log(f"version: {version}")
                        return True
            return False 
        except Exception as e:
            self.print_and_log(str(e))
            self.print_and_log(traceback.format_exc())
            self.print_and_log("Error downloading the alternative version","error")
            self.send_mail(f"Falló la descarga de otra versión del driver, por favor revise manualmente. Error{str(e)}",
                            "Error en descarga de Driver alternativo a version actual"
                            )
            return False


    def print_and_log(self, message, log_type="debug"):
        """
        Method to print and log the information, debug and error at the same time.

            Params: 
                message (str): message to print and log
                log_type (str): type of log to record, values could be: debug, info, warning, error.
                Default value is debug.
            Returns
                None
        """
        print(f"{log_type.upper()}: {message}")
        if log_type == "info":
            self.logger.info(message)
        elif log_type == "warning":
            self.logger.warning(message)
        elif log_type == "error":
            self.logger.error(message)
        else: 
            self.logger.debug(message)

    
    

    def send_mail(self, message, subject):
        html = f"<h2>Error chequeo de driver</h2>\
                <p>{message}</p>\
                <p>Atte: Bot de chequeo de driver de respaldo</p>\
                "
        try:
            credentials = (self.CLIENT_ID, str(self.token_mail))
            account = Account(credentials, auth_flow_type="credentials", tenant_id=self.TENANT_ID)
            token_backend = FileSystemTokenBackend(token_path=self.CURRENT_PATH, token_filename="o365_token.txt")
            account = Account(credentials, token_backend=token_backend)
            if not account.is_authenticated:
                account = Account(credentials, auth_flow_type="credentials", tenant_id=self.TENANT_ID)
                account.authenticate()
                self.print_and_log("Authenticated")
            my_protocol = MSGraphProtocol("beta", str(self.user_mail))
            account = Account(credentials, protocol=my_protocol)
            m = account.new_message()
            m.to.add(self.EMAILS_TO_REPORT)
            m.subject = f"{subject} Fecha: {self.today_datetime.strftime('%d-%m-%Y %H:%M')}"
            m.body = html
            m.send()
        except Exception as e:
            self.print_and_log(str(e))
            self.print_and_log(traceback.format_exc())


    def import_data(self):
        """
        Method to get the properties, such as BASE_URL for drivers resourses, driver name, etc.

            Params: 
                None
            Returns 
                None
        """
        try:
            with open(self.PATH_OF_PROPERTIES) as f:
                data = json.load(f)
            return data
        except Exception as e:
            self.print_and_log("Log error of version, notiftying via mail...") #notify that there is an error, probably have to solve manually 
            self.send_mail(f"No se pudo descargar ninguna version del driver, por favor, reviselo manualmente. Error: {str(e)}",
                            "Error cargando archivo de propiedades"
                            )
            

    def copy_to_path(self):
        """
        Method to copy downloaded driver (.exe) to the other folders of the bots.
            
            Params: 
                None
            Returns: 
                None
        """
        for folder in self.BOT_FOLDERS:
            shutil.copy2(self.DRIVER_NAME, folder)
        self.print_and_log(f"Copied to folders: {self.BOT_FOLDERS}")

if __name__=="__main__":
    s = EdgeDriverLocal()
    if not s.test():
        s.download() # try the main version
        if not s.test(): # if doesn't work, try other versions
            status = s.try_another_version()
            if not status:
                s.print_and_log("Log error of version, notiftying via mail...") #notify that there is an error, probably have to solve manually 
                s.send_mail("No se pudo descargar ninguna version del driver, por favor, reviselo manualmente. ",
                            "Error en driver"
                            )
                exit()
            else: # another version works, so let's copy to the others folder
                s.copy_to_path()
        else: # downloaded version works, so let's copy to the others folder
            s.copy_to_path()
    s.print_and_log("Driver working, all right") #log the message
