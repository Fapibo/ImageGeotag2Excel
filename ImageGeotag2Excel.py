from tkinter import filedialog, Tk, StringVar, Label, Button
import os
from GPSPhoto import gpsphoto
import xlsxwriter
import datetime

class Image2GeotagGUI:
    def __init__(self, master):
        self.GeoDataList = []
        self.master = master
        master.title("ImageGeotag2Excel")
        master.minsize(300, 50)
        
        self.folder_path = StringVar()
        self.Info = StringVar()
        
        self.Infolabel = Label(master, textvariable=self.Info)
        self.Infolabel.grid(row=0, column=1)
        self.InfoMsg("\"Browse...\" to choose source folder.", "Info")

        self.Captionlabel = Label(master, text="Image Folder:")
        self.Captionlabel.grid(row=1, column=0)
        
        self.buttonBrowse = Button(master, text="Browse...", command=self.BBrowse)
        self.buttonBrowse.grid(row=1, column=2)
 
        self.Pathlabel = Label(master, textvariable=self.folder_path)
        self.Pathlabel.grid(row=1, column=1)
        
        self.Go_button = Button(master, text="Go!", command=lambda: self.BGo(self.folder_path.get()))
        self.Go_button.grid(row=2, column=1)

        self.close_button = Button(master, text="Exit", command=master.quit)
        self.close_button.grid(row=2, column=2)

        

    def BBrowse(self):
        self.folder_path.set(filedialog.askdirectory())
        self.InfoMsg("Clic Go to launch images analysis.", "Info")

    def BGo(self , foldPath):
        self.GeoDataList = []
        valid_images = [".jpg",".jpeg"]
        try:
            FileList = os.listdir(foldPath)
        except:
            self.InfoMsg("Cannot read folder: " + foldPath, "Error")
        else: 
            for f in FileList:
                ext = os.path.splitext(f)[1]
                if ext.lower() in valid_images:
                    self.InfoMsg("Process: " + f, "Info")
                    GeoData = gpsphoto.getGPSData(foldPath + '/' + f)
                    self.GeoDataList.append([foldPath, f, GeoData['Latitude'], GeoData['Longitude'], GeoData['Altitude']])
            #print(self.GeoDataList)
            if len(self.GeoDataList) != 0:
                self.ExcelExport(foldPath)
            else:
                self.InfoMsg("No pictures found in folder " + foldPath, "Error")
            
    def InfoMsg(self,Message,Type):
        self.Info.set(Message)
        if Type == "Error":
            self.Infolabel.configure(bg="red")
        elif Type == "Success":  
            self.Infolabel.configure(bg="PaleGreen1")
        else:
            self.Infolabel.configure(bg="white")
        self.Infolabel.update()
        
    def ExcelExport(self, foldPath):     
        XlsName = foldPath + '/' + datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S ") + 'ImageGeotag2Excel.xlsx'
        # Create a workbook and add a worksheet.
        try:
            workbook = xlsxwriter.Workbook(XlsName)
            worksheet = workbook.add_worksheet()
        except:
            self.InfoMsg("Cannot write in folder: " + foldPath, "Error")
        else:    
            ColName=0
            ColLat=1
            ColLong=2
            ColAlt=3
            ColGoogle=4
            ColOSM=5
            ColGeoport=6
            
			# Format  excel
            worksheet.write(0, ColName, "Name")
            worksheet.write(0, ColLat, "Lat")
            worksheet.write(0, ColLong, "Long")
            worksheet.write(0, ColAlt, "Alt")   

            worksheet.set_column(ColName, ColAlt, 15)
            worksheet.set_column(ColGoogle, ColGeoport, 25)
                
            cell_format = workbook.add_format({'bold': True})
            worksheet.set_row(0, None, cell_format)   
            
            for i in range(len(self.GeoDataList)):
                self.InfoMsg("Prepare excel: " + str(i/len(self.GeoDataList)) +"%", "Info")
                JPGName=self.GeoDataList[i][1]
                worksheet.write_url(i+1, ColName, JPGName, string=JPGName)
                worksheet.write(i+1, ColLat, self.GeoDataList[i][2])
                worksheet.write(i+1, ColLong, self.GeoDataList[i][3])
                worksheet.write(i+1, ColAlt, self.GeoDataList[i][4])
                MapsURL = "https://www.google.com/maps/search/?api=1&query="
                MapsURL += str(self.GeoDataList[i][2]) + "%2C"
                MapsURL += str(self.GeoDataList[i][3]) + "&basemap=satellite"
                worksheet.write_url(i+1, ColGoogle, MapsURL , string="Google Maps "+JPGName)
                
                MapsURL = "https://www.openstreetmap.org/#map=14/"
                MapsURL += str(self.GeoDataList[i][2]) + "/" + str(self.GeoDataList[i][3])
                worksheet.write_url(i+1, ColOSM, MapsURL , string="OpenStreetMaps "+JPGName)
                
                MapsURL = "https://www.geoportail.gouv.fr/carte?c="
                MapsURL += str(self.GeoDataList[i][3]) + "," + str(self.GeoDataList[i][2])
                MapsURL += "&z=15&l0=ORTHOIMAGERY.ORTHOPHOTOS::GEOPORTAIL:OGC:WMTS(1)&permalink=yes"
                worksheet.write_url(i+1, ColGeoport, MapsURL , string="Geoportail "+JPGName)
                
                #print(JPGName)             
            self.InfoMsg("Write " + XlsName, "Info")
            workbook.close()
            self.InfoMsg("Successful ! Excel written in folder: #" + foldPath + "#", "Success")
       
                
root = Tk()
my_gui = Image2GeotagGUI(root)
root.mainloop()