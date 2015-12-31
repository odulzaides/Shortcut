import win32com.client
import winshell
import os

class Skytap_file(object):
    
    def __init__(self,):
        self.file_1 = raw_input("Enter SkyTap file to use: ")
        self.folder  = raw_input("Enter course folder to create: ")
        # add file extension
        self.file_1 = self.file_1+'.txt'
        #  Create folder to pu the shortcuts in
        if not os.path.exists(self.folder): os.makedirs(self.folder)

    def fileOpen(self):
        
        self.skytap = []
        with open(self.file_1, 'r') as f:   
            for line in f:
                self.skytap.append(line.strip('\t\r\n'))
            #for i in range(18):
                #self.skytap.pop()
       
        return self.skytap        
    
    def shortcuts(self):

        for i in range(len(self.skytap)):
            self.path = r"{}\Host{}.url".format(self.folder,i)
            ws =win32com.client.Dispatch("wscript.shell")
            shortcut = ws.CreateShortcut(self.path)
            shortcut.TargetPath= self.skytap[i]
            shortcut.Save()
            copy_bat = r'{}\copy.bat'.format(self.folder)
            with open(copy_bat, 'w') as f:
                for i in range(1, len(c.skytap)):
                    f.write(r'copy host{0}.url \\10.10.10.{0}\c$\Users\user\Desktop'.format(i))
                    f.write('\n')

c = Skytap_file()
c.fileOpen()
c.shortcuts()