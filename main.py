#install openpyxl
import tkinter as tk
import os
import shutil #allow to copy and paste
from openpyxl import Workbook,load_workbook
from breezypythongui import EasyFrame


excel_File_Name = "BTRAN Wafer Map 93361-01.xlsx"
def main():
    root = tk.Tk()
    WINDOW(root)
    root.mainloop() 
    
    
#----------- This must be converted into a class --------------#


class WINDOW(EasyFrame,tk.Tk):
    
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
        self.user_frame = tk.LabelFrame(self.frame, text= "Select the file name")
        self.user_frame.grid(row=0,column=0) 
        
        self.boolVar = tk.BooleanVar() #check if the wafer is SemeFab or not
        SemeFab_option = tk.Checkbutton(self.user_frame,
                            text="Click here if the wafer is type SemeFab",
                            variable=self.boolVar,
                            onvalue=True,
                            offvalue=False)
        SemeFab_option.grid(row=2,column=0)
        
        self.ProberCondition = tk.BooleanVar()#check button to see if the prober is good or not 
        self.ProberState = tk.Checkbutton(self.user_frame,
                                    text="Does the probe work?",
                                    variable=self.ProberCondition,
                                    onvalue=True,
                                    offvalue=False,
                                    command=self.check)    
        self.ProberState.grid(row=3,column=0)
        self.ProberCondition_flag = self.ProberCondition.get()
        self.ProberCondition_entry = tk.Entry(self.user_frame)

        self.user_button= tk.Button(self.user_frame,text="print",command= self.ShowFinalWafer)
        self.user_button.grid(row= 4,column=0)

        self.file_label = tk.Label(self.user_frame,text="File name: ")
        self.file_label.grid(row=0,column=0)

        self.file_entry = tk.Entry(self.user_frame)
        self.file_entry.grid(row=0,column=1)
        print(self.file_entry.get())

    def pick_data(self,filename):
        file = open(filename , 'r')
        reads = file.read()
        Lines = reads.split()
        #print(Lines)
        i = 1
        number = [0,0]
        data = {}
        for Line in Lines:
            splitLine = Line.split(",")
            if(f"V_Base@frontside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                #print(f"die {i}")
                #print(splitLine)
                number[0] = float(splitLine[1])
                
            elif(f"V_Collector@Backside@HOME[{self.MAX_VOLTAGE_FOR_TEST}]" in Line):
                #print(splitLine)
                number[1] = float(splitLine[1])
                data[f"Die {i}"] = list(number)
                #print(number,"\n")
                i+=1  
        return(data)

    def die_text_box(self,row,column,die_Num,data,root):
        
        row_distance = row+self.cell_height*row*3
        column_distance = column+self.cell_width*column
        die_list = list(data.items())
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
        for datas in data.items():
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
                        self.die_text_box(row,column,i,data,root)
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
                        self.die_text_box(row,column,i,data,root)
                    i+=1
    #-----------------------------------------------------------------------#
    def extra_window(self,data,flag):
        root = tk.Toplevel()
        root.title("Wafermap")
        if(flag == True):
            self.display_wafer_SemeFab(root,data)
        else:
            self.display_wafer(root,data)

    def ShowFinalWafer(self):######################################################
        Name = str(self.file_entry.get())+'.kdf'
        """Second_Name = str()+'kdf' """
        print(Name)
        wafer_flag = self.boolVar.get()
        data = self.pick_data(Name)
        self.print_data(data)
        #---------------This Section is used to create the final excel File------
        """ProberCondition_flag = ProberCondition.get()
        if(ProberCondition_flag == True):
            #modify data change
        elif(ProberCondition_flag == False):
            #"""
        #create a new folder
        FolderName = str(self.file_entry.get())
        try:
            os.mkdir(FolderName)
        except:
            print(f"The Folder {FolderName} already exist")
        current_path = str(os.getcwd())
        destination_path = current_path +'\\'+ FolderName
        #shutil.copy(current_path+'\\'+excel_File_Name,destination_path)
        os.chdir(f"{current_path}\{FolderName}") #change the directory to save the excel Folder

        #filling The excel data with the dictionary that contents the information
        book = load_workbook(filename=f"{current_path}\{excel_File_Name}")
        sheet = book.active #current and ONLY sheet
        
        #make a for llop that fill the data with the dictionary data fetched from the .kdf file
        
        #create an array with the alphabet letters
        sheet_column = list(map(chr, range(ord('A'), ord('J')+1)))
        for column in sheet_column:
            for row in range(6,36):
                if("Die" in str(sheet[str(column)+str(row)].value)):
                    sheet[str(column)+str(row-1)] = data[sheet[str(column)+str(row)].value][0]
                    sheet[str(column)+str(row+1)] = data[sheet[str(column)+str(row)].value][1]
        book.save(filename = "BTRAN Wafer Map "+str(self.file_entry.get())+".xlsx")  
        os.chdir(f"{current_path}") #go back to regular folder
        #------------------------------------------------------------------------
        
        self.extra_window(data,wafer_flag)

        
    def check(self):
        ProberCondition_flag = self.ProberCondition.get()
        #print(ProberCondition_flag)
        
        if(ProberCondition_flag == True):
            ProberCondition_entry = tk.Entry(self.user_frame)
            ProberCondition_entry.grid(row=3,column=1)


if __name__ == "__main__":
    main()
    #----------------------------------------------------
    
    