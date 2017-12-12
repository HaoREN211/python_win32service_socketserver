import os
import os.path
import sys
import urllib.request
from urllib.parse import unquote
 
import pythoncom
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import threading
import time
from http.server import BaseHTTPRequestHandler


import socketserver 
import re
from socketserver import StreamRequestHandler as SRH 
from time import ctime 
import shutil

  
#如果自己要实现TCPServer,我们还要自己定义一个请求处理类，而且要继承BaseRequestHandler或者StreamRequestHandler，并且覆盖他的handle()处理函数，用来处理请求。
#Si on veut traiter les demandes avec type TCP, on doit définir une classe qui sert à traiter les demandes et qui hérite la classe BaseRequestHandler ou StreamRequestHandler. Après on doit aussi rédifinir la fonction handle() pour définir les traitements de demandes.
#If you want to process requests with TCP type, you must define a class that is used to process requests and that inherits the class BaseRequestHandler or StreamRequestHandler. After that we must also redefine the handle () function to define the request processing.

class Servers(SRH): 
  def handle(self): 
      #接收数据
      #reçcoit les données transférées par les clients
	  #receives data transferred by customers
      data = self.request.recv(1024) 

      #将接收的数据进行解码
      #Décodage des données que l'on a reçues
	  #Decoding the data we received
      data = data.decode("utf-8")
	  
      #将收到的数据存在Params这个数组里面
      #Stokage des données dans le tableau Params
	  #Storing data in the Params table
      Params={}
	  
      #初始化数组中urlRecuperation和statut对应的值
      #Initialisation des données pour les variables "urlRecuperation" et "statut"
	  #Initialization of the data
      Params["urlRecuperation"] = ""
      Params["statut"] = ""
      Params["urlConfirmation"] = ""
      URL_confirmation = ""
      projetId = ""
      #将所有接收到的提示消息存在下面指定的目录里
      #Sauvegarder l'historique de la notification reçuées
	  #Save notification history received
      data_with_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())+"\n"+data+"\n"
      f = open('C:/Historique_échange_notification.txt', 'a')
      f.write(data_with_time)
      f.close()
	  
      #取回只包含有客户传送数据的那行数据，取消掉header等等。支持GET和POST
      #Récupération que des données transférées par les clients(suppression de données de header et etc. supporter le type de transaction GET et POST)
	  #Recovering only data transferred by the clients (deleting header data and etc. supporting transaction type GET and POST
      res="";
      
      if data!='':
        for Input in data.split('\n'):
          if "urlRecuperation" in Input:
            res=Input
      if res!="":
        for Input in res.split("&"):
            (Command, Value)=Input.split('=', 1)
            if "siteWebId" in Command:
              Command = "siteWebId"
            if "projetId" in Command:
              Command = "projetId"
              projetId = Value
            if "ponctuel" in Command:
              Command = "ponctuel"
            if "urlRecuperation" in Command:
              Command = "urlRecuperation"
              Value = unquote(Value)
            if "reinitialisation" in Command:
              Command = "reinitialisation"
            if "statut " in Command:
              Command = "statut "
            if "urlConfirmation" in Command:
              Command = "urlConfirmation"
              Value = unquote(Value)
              URL_confirmation = unquote(Value)
            Params[Command]=Value
      if "SUCCESS" in Params["statut"]:
        #取回zip文件并进行解压
        #Récupération de fichier zip et le décompresser
		#Recovery the zip file and unzip it
        if Params["urlRecuperation"] != "":
          reponse_client = """
HTTP/1.1 200 OK
"""
          self.request.sendall(bytes(reponse_client, "utf-8"))
          self.request.close()
          re_words=re.compile("[0-9]{8}")  
          m =  re_words.search(Params["urlRecuperation"],0)
          file_split = Params["urlRecuperation"].split("/")
          num_split = len(file_split)
          FileDestZip = file_split[num_split-1]
          if "true" in Params["reinitialisation"]:
            tmp_name_split = FileDestZip.split(".")
            FileDestZip = tmp_name_split[0]+"_COMPLET.zip"
          FileDestZip = "C:/VD/Connecteurs/sit/in/apidae/Export/"+FileDestZip
          FileDest="C:/VD/Connecteurs/sit/in/data/ADT07/APT09_"+m.group()+"_VDV_1_0_0"
          UrlFile=urllib.request.urlopen(Params["urlRecuperation"])
          SitraFile=open(FileDestZip, "wb")
          SitraFile.write(UrlFile.read())
          SitraFile.close()  
          #shutil.unpack_archive(FileDestZip, FileDest)
        UrlConfirmation = urllib.request.urlopen(URL_confirmation)
        contenu_confirmation = UrlConfirmation.read()
        contenu_confirmation = contenu_confirmation + "\n--------------------------------------------------------------------\n"
        f = open('C:/VD/Connecteurs/sit/in/arch/ADT07/Historique_échange_notification.txt', 'a')
        f.write(contenu_confirmation)
        f.close()
        if( UrlConfirmation.status != 200 ):  
          exit() 
            # os.system("C:/sitra_import/teste_1/echoHello.bat")
        curry_time_string = time.strftime("%Y%m%d", time.localtime())
        file_url_confirmation_dir = "C:/VD/Connecteurs/sit/in/apidae/Export/"
        urlConfirmation_file=open(file_url_confirmation_dir+"apidae_url_confirmation_"+curry_time_string+"_"+projetId+".txt", "w")
        urlConfirmation_file.write(Params["urlConfirmation"])
        urlConfirmation_file.close()  
		

  
class PythonService(win32serviceutil.ServiceFramework):   
    #服务名  
    #Nom de service
    _svc_name_ = "rec_notifi" 	
    #服务显示名称  
    #Le nom de service affiché
    _svc_display_name_ = "rec_notifi"  
    #服务描述  
    #La description de service
    _svc_description_ = "{La description de votre service}"  
  
    def __init__(self, args):   
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None,0,0,None)
        self.isAlive = True
        socket.setdefaulttimeout(60)
        self.NeedStop=False
        path = os.path.dirname(__file__)	
        # host = 'localhost'
        host = "{Host}"
        port = 8011
        addr = (host,port)
        self.server = socketserver.ThreadingTCPServer(addr,Servers)
       
    def SvcDoRun(self): 	 
        self.server.serve_forever()		
        while self.isAlive:  
            time.sleep(1)  	
        
    def SvcStop(self): 
        self.server.shutdown() 	 
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)   
        win32event.SetEvent(self.hWaitStop)   
        self.isAlive = False  
  
if __name__=='__main__':   
    win32serviceutil.HandleCommandLine(PythonService)