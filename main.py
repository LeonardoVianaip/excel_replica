import tkinter as tk
cell_width=8
cell_height=1

filename = '93361-07.kdf'
def main():
    file_label = tk.Label(user_frame,text="File name: ")
    file_label.grid(row=0,column=0)

    
    
    
    #extra_window_flag = False
    
    """if(SelectWafer):
        display_wafer_SemeFab(frame,data)
    else:
        display_wafer(frame,data)"""
    print(file_entry.get())
    
    
#----------- This must be converted into a class --------------#

def pick_data(filename):
    file = open(filename , 'r')
    reads = file.read()
    Lines = reads.split()
    #print(Lines)
    i = 1
    number = [0,0]
    data = {}
    for Line in Lines:
        splitLine = Line.split(",")
        if("V_Base@frontside@HOME[100]" in Line):
            #print(f"die {i}")
            #print(splitLine)
            number[0] = float(splitLine[1])
            
        elif("V_Collector@Backside@HOME[100]" in Line):
            #print(splitLine)
            number[1] = float(splitLine[1])
            data[f"die {i}"] = list(number)
            #print(number,"\n")
            i+=1  
    return(data)

def die_text_box(row,column,die_Num,data,root):
    
    row_distance = row+cell_height*row*3
    column_distance = column+cell_width*column
    die_list = list(data.items())
    die_number = die_Num
    
    
    die = die_list[die_number][0]
    die_value = die_list[die_number][1]
    die_value_top = round(die_value[0],2)
    die_value_bottom = round(die_value[1],2)
    
    color = ''
    if(die_value_top > 45 and die_value_bottom >45):
        color = 'green'
    elif((die_value_top > 45 and die_value_bottom <45)or(die_value_top < 45 and die_value_bottom >45)):
        color = 'orange'
    else:
        color = 'red'

    display_top_value = tk.Label(root,text=die_value_top,width=cell_width,height=cell_height,borderwidth=1,relief='solid',background=color)
    display_top_value.grid(row=row_distance,column=column_distance)
    
    display_die_number = tk.Label(root,text=die,width=cell_width,height=cell_height,borderwidth=1,relief='solid',background=color)
    display_die_number.grid(row=row_distance+1,column=column_distance)
    
    display_bottom_value = tk.Label(root,text=die_value_bottom,width=cell_width,height=cell_height,borderwidth=1,relief='solid',background=color)
    display_bottom_value.grid(row=row_distance+2,column=column_distance)
    
    #print(die,die_value_top,die_value_bottom)
    
def blank_box(row,column,root):
    row_distance = row+cell_height*row*3
    column_distance = column+cell_width*column
    blank =tk.Label(root,width=cell_width,height=cell_height,borderwidth=0)
    blank.grid(row=row_distance,column=column_distance)
    
    
def print_data(data):
    print("('die #', [frontside[V], backside[V]])")
    for datas in data.items():
        print(datas)

def display_wafer(root,data):
    i = 0
    for row in range(10):
        for column in range(10):
            if(row == 0 and (column < 2 or column > 7)):
                blank_box(row=row,column=column,root=root)
            elif(row == 1 and (column < 1 or column > 8)):
                blank_box(row=row,column=column,root=root)
            elif(row == 8 and (column < 1 or column > 8)):
                blank_box(row=row,column=column,root=root)
            elif(row == 9 and (column < 2 or column > 7)):
                blank_box(row=row,column=column,root=root)
            else:
                if(i<88):
                    die_text_box(row,column,i,data,root)
                i+=1

def display_wafer_SemeFab(root,data):
    i = 0
    for row in range(12):
        for column in range(12):
            if(row == 0 and (column < 4 or column > 7)):
                blank_box(row=row,column=column,root=root)
            elif(row == 1 and (column < 2 or column > 9)):
                blank_box(row=row,column=column,root=root)
            elif((row == 2 or row == 3) and (column < 1 or column > 10)):
                blank_box(row=row,column=column,root=root)
            elif((row == 4 or row == 5) and (column < 0 or column > 11)):
                blank_box(row=row,column=column,root=root)
            elif((row == 6 or row == 7) and (column < 1 or column > 10)):
                blank_box(row=row,column=column,root=root)
            elif(row == 8 and (column < 2 or column > 9)):
                blank_box(row=row,column=column,root=root)
            elif(row == 9 and (column < 4 or column > 7)):
                blank_box(row=row,column=column,root=root)
            else:
                if(i<88):
                    die_text_box(row,column,i,data,root)
                i+=1
                



#-----------------------------------------------------------------------#
def extra_window(data,flag):
    root = tk.Toplevel()
    root.title("Wafermap")
    if(flag == True):
        display_wafer_SemeFab(root,data)
    else:
        display_wafer(root,data)

def ShowFinalWafer():######################################################
    Name = str(file_entry.get())+'.kdf'
    print(Name)
    flag = boolVar.get()
    print(flag)
    
    data = pick_data(Name)
    print_data(data)
    extra_window(data,flag)
    
    
if __name__ == "__main__":
    root = tk.Tk()
    """ GUI to short the data evaluation """
    root.title("User Setup")
    #root.geometry('640x480')
    #--------------------------------------------
    frame = tk.Frame(root)
    frame.pack() 
    
    user_frame = tk.LabelFrame(frame, text= "Select the file name")
    user_frame.grid(row=0,column=0) 
    
    file_entry = tk.Entry(user_frame)
    file_entry.grid(row=0,column=1)
    
    user_button= tk.Button(user_frame,text="print",command= ShowFinalWafer)
    user_button.grid(row= 4,column=0)
    
    boolVar = tk.BooleanVar() #check if the wafer is SemeFab or not
    SemeFab_option = tk.Checkbutton(user_frame,
                        text="Click here if the wafer is type SemeFab",
                        variable=boolVar,
                        onvalue=True,
                        offvalue=False)
    SemeFab_option.grid(row=2,column=0)
    
    ProberState = tk.Checkbutton(user_frame,
                        text="Does the probe work?")
    ProberState.grid(row=3,column=0)
    main()
    #----------------------------------------------------
    root.mainloop() 
