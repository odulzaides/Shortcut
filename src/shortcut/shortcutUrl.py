import win32com.client
import winshell
import os


'''
Created on Apr 9, 2015

@author: user
'''
class Skytap_file(object):
    
    def __init__(self,):
        self.file_1 = raw_input("Enter SkyTap file to use: ")
        self.input  = raw_input("Enter course folder to create: ")
        self.file_1 = self.file_1+'.txt'
        self.folder = "C:\Users\user\Desktop"
        self.folder = os.path.join(self.folder, self.input)
        if not os.path.exists(self.folder): os.makedirs(self.folder)

    def fileOpen(self):
        
        self.skytap = []
        with open(self.file_1, 'r') as f:   
            for line in f:
                self.skytap.append(line.strip('\t\n'))
        return self.skytap        
        
    def createHosts(self):
        
        self.hosts = ['host{}'.format(i+1) for i in range(1+len(self.skytap))]
        return self.hosts
#       print hosts

    
    def shortcuts(self):
         
        self.path = r"{}\Host{}.url".format(self.folder,1)
        ws =win32com.client.Dispatch("wscript.shell")
        shortcut = ws.CreateShortcut(self.path)
        shortcut.TargetPath= self.skytap[0]
        shortcut.Save()
        
        


        
           
c = Skytap_file()
c.fileOpen()
c.createHosts()
c.skytap
c.shortcuts()


