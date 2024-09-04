#install openpyxl
import tkinter as tk
from tkinter import filedialog
import os
import os.path
import shutil #allow to copy and paste
from openpyxl import Workbook,load_workbook


excel_File_Name = "Blank - 6in Wafer Map - AFSW.xlsm"
def main():
    root = tk.Tk()
    WINDOW(root)
    root.mainloop() 
    
    
#----------- This must be converted into a class --------------#


class WINDOW(tk.Tk):
    
    def __init__(self,root):
        self.VOLTAGE_RATE = 45
        self.MAX_VOLTAGE_FOR_TEST = 100
        self.cell_width=8
        self.cell_height=1
        self.filename = ''
        self.root = root
        """ GUI to short the data evaluation """
        self.root.title("User Setup")
        #root.geometry('640x480')
        self.frame = tk.Frame(self.root)
        self.frame.pack() 
        
        
        #--------------------------------------------
        self.user_frame = tk.LabelFrame(self.frame, text= "File name")
        self.user_frame.grid(row=1,column=0) 
        
        self.folder_name_label = tk.Label(self.user_frame,text = "select your folder or directory")
        self.folder_name_label.grid(row = 0, column = 0)
        self.choose_folder_flag = False
        self.choose_folder_button= tk.Button(self.user_frame,
                                             text="Browse",
                                             command= self.change_folder)
        self.choose_folder_button.grid(row= 0,column=1)
        
        self.directory = self.choose_folder(False)
        
        self.current_path = tk.Label(self.frame, text= f"Current directory: {self.directory}")
        self.current_path.grid(row=0,column=0) 
        
        self.file_entry = tk.Entry(self.user_frame)
        self.file_entry.grid(row=1,column=1)
        print(self.file_entry.get())
        
        self.boolVar = tk.BooleanVar() #check if the wafer is SemeFab or not
        SemeFab_option = tk.Checkbutton(self.user_frame,
                            text="Click here if the wafer is type SemeFab",
                            variable=self.boolVar,
                            onvalue=True,
                            offvalue=False)
        SemeFab_option.grid(row=2,column=0)


        self.excel_flag = tk.BooleanVar() #chek if the user wish to add an excel folder
        SemeFab_option = tk.Checkbutton(self.user_frame,
                            text="Add excel file?",
                            variable=self.excel_flag,
                            onvalue=True,
                            offvalue=False)
        SemeFab_option.grid(row=2,column=1)
        
        self.ProberCondition = tk.BooleanVar()#check button to see if the prober is good or not 
        self.ProberState = tk.Checkbutton(self.user_frame,
                                    text="Check here if only one side of the Prober works",
                                    variable=self.ProberCondition,
                                    onvalue=True,
                                    offvalue=False,
                                    command=self.check)    
        self.ProberState.grid(row=3,column=0)
        self.goodSide = tk.BooleanVar()

        self.ProberCondition_flag = self.ProberCondition.get()
        self.ProberCondition_entry = tk.Entry(self.user_frame)
        self.ProberCondition_entry.grid(row=3,column=1)
        self.ProberCondition_entry.config(state='disabled')

        self.user_button= tk.Button(self.user_frame,text="print",command= self.ShowFinalWafer)
        self.user_button.grid(row= 5,column=0)
        self.root.bind('<Return>', self.ShowFinalWafer_Enter)
        
        
        self.file_label = tk.Label(self.user_frame,text="File name: ")
        self.file_label.grid(row=1,column=0)
    
    def change_folder(self):
        self.directory = self.choose_folder(True)
        print(self.directory)
        self.choose_folder_flag = False
        
    def choose_folder(self,flag):
        self.choose_folder_flag = flag
        if(self.choose_folder_flag == True):
            self.current_path.destroy()
            self.directory = filedialog.askdirectory()
            self.current_path = tk.Label(self.frame, text= f"Current directory: {self.directory}")
            self.current_path.grid(row=0,column=0) 
        else:
            self.directory = os.getcwd()
        return self.directory
        

    def pick_data(self,filename):
        if(self.ProberCondition.get() == True):
            fileLocation = self.directory
            file1 = open(f"{fileLocation}\\{filename}" , 'r')
            reads1 = file1.read()
            Lines1 = reads1.split()

            file2 = open(f"{fileLocation}\\{self.ProberCondition_entry.get()}.kdf" , 'r')
            reads2 = file2.read()
            Lines2 = reads2.split()

            i = 1
            number1 = [0,0]
            number2 = [0,0]
            self.data = {}
            data1 = {}
            data2 = {}
            
            for Line in Lines1:
                splitLine = Line.split(",")
                if(f"V_Base@frontside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                    number1[0] = float(splitLine[1])
                    
                elif(f"V_Collector@Backside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                    number1[1] = float(splitLine[1])
                    data1[f"Die {i}"] = list(number1)#list([i,i]) #use this line for troubleshooting
                    i+=1

            i = 1

            for Line in Lines2:
                splitLine = Line.split(",")
                if(f"V_Base@frontside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                    number2[0] = float(splitLine[1])
                    
                elif(f"V_Collector@Backside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                    number2[1] = float(splitLine[1])
                    data2[f"Die {i}"] = list(number2)#list([i,i]) #use this line for troubleshooting
                    i+=1
           
            i = 1
            row = 1
            j = 0
            ##################Sample for future condition######################
            if(self.goodSide.get() == True): #Wafer flipped to the TOP!
                print("Top is Selected")
                i = 1
                row = 1
                j = 6
                for datas in data1:
                    if(self.boolVar.get() == False):
                        if(row == 1 and i == 1):
                            j = 6
                        if(row == 2 and i == 7):
                            j = 14
                        if(row == 3 and i == 15):
                            j = 24
                        if(row == 4 and i == 25):
                            j = 34
                        if(row == 5 and i == 35):
                            j = 44
                        if(row == 6 and i == 45):
                            j = 54
                        if(row == 7 and i == 55):
                            j = 64
                        if(row == 8 and i == 65):
                            j = 74
                        if(row == 9 and i == 75):
                            j = 82
                        if(row == 10 and i == 83):
                            j = 88
                    else:
                        if(row == 1 and i == 1):
                            j = 4
                        if(row == 2 and i == 5):
                            j = 12
                        if(row == 3 and i == 13):
                            j = 22
                        if(row == 4 and i == 23):
                            j = 32
                        if(row == 5 and i == 33):
                            j = 44
                        if(row == 6 and i == 45):
                            j = 56
                        if(row == 7 and i == 57):
                            j = 66
                        if(row == 8 and i == 67):
                            j = 76
                        if(row == 9 and i == 77):
                            j = 84
                        if(row == 10 and i == 85):
                            j = 88
                    self.data[f"Die {i}"] = list([data1[f"Die {i}"][0],data2[f"Die {j}"][0]])
                    i+=1
                    j-=1
                    
                    if(i == 4):
                        row = row+1
                    elif(i == 12):
                        row = row+1
                    elif(i == 22):
                        row = row+1
                    elif(i == 32):
                        row = row+1
                    elif(i == 44):
                        row = row+1
                    elif(i == 56):
                        row = row+1
                    elif(i == 66):
                        row = row+1
                    elif(i == 76):
                        row = row+1
                    elif(i == 84):
                        row = row+1
                    elif(i == 88):
                        row = row+1

            elif(self.goodSide.get() == False):
                print("Bottom is selected")
                i = 1
                row = 1
                j = 6
                for datas in data1:
                    if(self.boolVar.get() == False):
                        if(row == 1 and i == 1):
                            j = 6
                        if(row == 2 and i == 7):
                            j = 14
                        if(row == 3 and i == 15):
                            j = 24
                        if(row == 4 and i == 25):
                            j = 34
                        if(row == 5 and i == 35):
                            j = 44
                        if(row == 6 and i == 45):
                            j = 54
                        if(row == 7 and i == 55):
                            j = 64
                        if(row == 8 and i == 65):
                            j = 74
                        if(row == 9 and i == 75):
                            j = 82
                        if(row == 10 and i == 83):
                            j = 88
                    else:
                        if(row == 1 and i == 1):
                            j = 4
                        if(row == 2 and i == 5):
                            j = 12
                        if(row == 3 and i == 13):
                            j = 22
                        if(row == 4 and i == 23):
                            j = 32
                        if(row == 5 and i == 33):
                            j = 44
                        if(row == 6 and i == 45):
                            j = 56
                        if(row == 7 and i == 57):
                            j = 66
                        if(row == 8 and i == 67):
                            j = 76
                        if(row == 9 and i == 77):
                            j = 84
                        if(row == 10 and i == 85):
                            j = 88
                    self.data[f"Die {i}"] = list([data1[f"Die {i}"][1],data2[f"Die {j}"][1]])
                    i+=1
                    j-=1
                    
                    if(i == 4):
                        row = row+1
                    elif(i == 12):
                        row = row+1
                    elif(i == 22):
                        row = row+1
                    elif(i == 32):
                        row = row+1
                    elif(i == 44):
                        row = row+1
                    elif(i == 56):
                        row = row+1
                    elif(i == 66):
                        row = row+1
                    elif(i == 76):
                        row = row+1
                    elif(i == 84):
                        row = row+1
                    elif(i == 88):
                        row = row+1
            ###################################################
      
        elif(self.ProberCondition.get() == False):
            fileLocation = self.directory
            file = open(f"{fileLocation}\\{filename}" , 'r')
            reads = file.read()
            Lines = reads.split()
            #print(Lines)
            i = 1
            number = [0,0]
            self.data = {}
            for Line in Lines:
                splitLine = Line.split(",")
                if(f"V_Base@frontside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                    #print(f"die {i}")
                    #print(splitLine)
                    number[0] = float(splitLine[1])
                    
                elif(f"V_Collector@Backside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                    #print(splitLine)
                    number[1] = float(splitLine[1])
                    self.data[f"Die {i}"] = list(number)
                    #print(number,"\n")
                    i+=1  
        

        return(self.data)

    def die_text_box(self,row,column,die_Num,data,root):
        
        row_distance = row+self.cell_height*row*3
        column_distance = column+self.cell_width*column
        die_list = list(self.data.items())
        die_number = die_Num
        
        
        die = die_list[die_number][0]
        die_value = die_list[die_number][1]
        die_value_top = round(die_value[0],2)
        die_value_bottom = round(die_value[1],2)
        
        color = ''
        if(die_value_top > self.VOLTAGE_RATE and die_value_bottom >self.VOLTAGE_RATE):
            color = 'green'
        elif((die_value_top > self.VOLTAGE_RATE/2 and die_value_bottom < self.VOLTAGE_RATE)or(die_value_top < self.VOLTAGE_RATE and die_value_bottom >self.VOLTAGE_RATE/2)):
            color = 'orange'
        else:
            color = 'white'

        display_top_value = tk.Label(root,text=die_value_top,width=self.cell_width,height=self.cell_height,borderwidth=1,relief='solid',background=color)
        display_top_value.grid(row=row_distance,column=column_distance)
        
        display_die_number = tk.Label(root,text=die,width=self.cell_width,height=self.cell_height,borderwidth=1,relief='solid',background=color)
        display_die_number.grid(row=row_distance+1,column=column_distance)
        
        display_bottom_value = tk.Label(root,text=die_value_bottom,width=self.cell_width,height=self.cell_height,borderwidth=1,relief='solid',background=color)
        display_bottom_value.grid(row=row_distance+2,column=column_distance)
        
        #print(die,die_value_top,die_value_bottom)
        
    def blank_box(self,row,column,root):
        row_distance = row+self.cell_height*row*3
        column_distance = column+self.cell_width*column
        blank =tk.Label(root,width=self.cell_width,height=self.cell_height,borderwidth=0)
        blank.grid(row=row_distance,column=column_distance)
        
        
    def print_data(self,data):
        print("('die #', [frontside[V], backside[V]])")
        for datas in self.data.items():
            print(datas)

    def display_wafer(self,root,data):
        i = 0
        for row in range(10):
            for column in range(10):
                if(row == 0 and (column < 2 or column > 7)):
                    self.blank_box(row=row,column=column,root=root)
                elif(row == 1 and (column < 1 or column > 8)):
                    self.blank_box(row=row,column=column,root=root)
                elif(row == 8 and (column < 1 or column > 8)):
                    self.blank_box(row=row,column=column,root=root)
                elif(row == 9 and (column < 2 or column > 7)):
                    self.blank_box(row=row,column=column,root=root)
                else:
                    if(i<88):
                        self.die_text_box(row,column,i,self.data,root)
                    i+=1

    def display_wafer_SemeFab(self,root,data):
        i = 0
        for row in range(12):
            for column in range(12):
                if(row == 0 and (column < 4 or column > 7)):
                    self.blank_box(row=row,column=column,root=root)
                elif(row == 1 and (column < 2 or column > 9)):
                    self.blank_box(row=row,column=column,root=root)
                elif((row == 2 or row == 3) and (column < 1 or column > 10)):
                    self.blank_box(row=row,column=column,root=root)
                elif((row == 4 or row == 5) and (column < 0 or column > 11)):
                    self.blank_box(row=row,column=column,root=root)
                elif((row == 6 or row == 7) and (column < 1 or column > 10)):
                    self.blank_box(row=row,column=column,root=root)
                elif(row == 8 and (column < 2 or column > 9)):
                    self.blank_box(row=row,column=column,root=root)
                elif(row == 9 and (column < 4 or column > 7)):
                    self.blank_box(row=row,column=column,root=root)
                else:
                    if(i<88):
                        self.die_text_box(row,column,i,self.data,root)
                    i+=1
    #-----------------------------------------------------------------------#
    def extra_window(self,data,flag):
        root = tk.Toplevel()
        root.title("Wafermap")
        if(flag == True):
            self.display_wafer_SemeFab(root,self.data)
        else:
            self.display_wafer(root,self.data)

    def ShowFinalWafer(self):######################################################
        Name = str(self.file_entry.get())+'.kdf'
        #Second_Name = str()+'kdf' 
        print(Name)
        wafer_flag = self.boolVar.get()
        self.data = self.pick_data(Name)
        self.print_data(self.data)
        
        self.extra_window(self.data,wafer_flag)

        if (self.excel_flag.get() == True):
            self.excel_file()
            

    def ShowFinalWafer_Enter(self,event):
        self.ShowFinalWafer()
        
    def E1E2read(self):
        
        if(os.path.exists(f"{self.directory}\\{self.file_entry.get()}_E1E2.kdf") and
           os.path.exists(f"{self.directory}\\{self.file_entry.get()}_E2E1.kdf")):
            
            print(f"Files {self.file_entry.get()}_E1E2 and {self.file_entry.get()}_E2E1 exist")
            #Create a function to fetch the data of e1e2 files and create and fill the excel folder
            
            FolderName = str(self.file_entry.get())
            
            filename = f"{self.file_entry.get()}_E1E2.kdf"
            fileLocation = self.directory
            E1E2_file = open(f"{fileLocation}\\{filename}" , 'r')
            E1E2_reads = E1E2_file.read()
            E1E2_Lines = E1E2_reads.split()

            E2E1_file = open(f"{fileLocation}\\{self.file_entry.get()}_E2E1.kdf" , 'r')
            E2E1_reads = E2E1_file.read()
            E2E1_Lines = E2E1_reads.split()

            i = 1
            E1E2_number = 0
            E2E1_number = 0
            E1E2_data = {}
            E2E1_data = {}
            
            for Line in E1E2_Lines:
                splitLine = Line.split(",")
                if(f"V_Pos@itm@HOME[1]" in Line):
                    E1E2_number = float(splitLine[1])
                    E1E2_data[f"Die {i}"] = E1E2_number
                    i+=1
            
            i = 1
            
            for Line in E2E1_Lines:
                splitLine = Line.split(",")
                if(f"V_Pos@itm@HOME[1]" in Line):
                    E2E1_number = float(splitLine[1])
                    E2E1_data[f"Die {i}"] = E2E1_number
                    i+=1
            
            print("E1E2 ('die #', Result)")
            for datas in E1E2_data.items():
                print(datas)
            
            
            print("E2E1 ('die #', Result)")
            for datas in E2E1_data.items():
                print(datas)
            
            
            current_path = str(os.getcwd())
            destination_path = current_path +'\\'+ FolderName


            
            self.sheet['K52'] = "aiura"
            
            for i in range(1,89):
                row = 51+i
                self.sheet[f"K{row}"] = E1E2_data[f"Die {i}"]
                self.sheet[f"L{row}"] = E2E1_data[f"Die {i}"]
            self.sheet['K52'] = "aiura"
            
            
        else:
            print(f"Files {self.file_entry.get()}_E1E2 and {self.file_entry.get()}_E1E2 are missing")
        
    
    def excel_file(self):
        ProberCondition_flag = self.ProberCondition.get()
        #create a new folder -------------------------------comment this Line!
        FolderName = str(self.file_entry.get())
        
        current_path = str(os.getcwd())
        destination_path = current_path +'\\'+ FolderName
        #os.chdir(f"{current_path}+\\+{FolderName}") #change the directory to save the excel Folder

        #Extract the lot number and the wafer number
        split_name = FolderName.split('-')
        
        #filling The excel data with the dictionary that contents the information
        book = load_workbook(filename=f"{current_path}\\{excel_File_Name}")
        self.sheet = book.active #current and ONLY sheet
        
        
        
        self.sheet['C2'] = split_name[0]
        self.sheet['E2'] = split_name[1]
        #create an array with the alphabet letters
        sheet_column = list(map(chr, range(ord('A'), ord('J')+1)))
        #make a for loop that fill the data with the dictionary data fetched from the .kdf file
        for column in sheet_column:
            for row in range(6,36):
                if("Die" in str(self.sheet[str(column)+str(row)].value)):
                    self.sheet[str(column)+str(row-1)] = self.data[self.sheet[str(column)+str(row)].value][0]
                    self.sheet[str(column)+str(row+1)] = self.data[self.sheet[str(column)+str(row)].value][1]
        self.E1E2read()
            
        book.save(filename = "6in Wafer Map - AFSW "+str(self.file_entry.get())+".xlsx")  
        book.close()
        
        
        
    def check(self):
        ProberCondition_flag = self.ProberCondition.get()
        #print(ProberCondition_flag)
        
        if(ProberCondition_flag == True):
            self.ProberCondition_entry.config(state='normal')
            
            self.RadioTop=tk.Radiobutton(self.user_frame, 
                                text="Top",
                                variable = self.goodSide,
                                value=True)
            self. RadioTop.grid(row = 4 ,column = 0)

            self.RadioBottom=tk.Radiobutton(self.user_frame, 
                                text="Button",
                                variable = self.goodSide,
                                value = False)
            self. RadioBottom.grid(row = 4 ,column = 1)
        else:
            self.ProberCondition_entry.config(state='disabled')
            self.RadioTop.config(state="disabled")
            self.RadioBottom.config(state="disabled")




if __name__ == "__main__":
    main()
    #----------------------------------------------------
    
    