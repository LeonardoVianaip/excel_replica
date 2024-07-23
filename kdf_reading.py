import tkinter as tk
def main():
    filename = "93361-11.kdf"
    data = pick_data(filename)
    print_data(data)
    
    root = tk.Tk()
    """scroll_bar = tk.Scrollbar(root) 
    scroll_bar.pack( side="right", 
                    fill = "y" ) 
    root.geometry("640x480")"""
    root.title("test")
    i = 0
    """for datas in data.items():"""
    #die_text_box(0,0,0,data)
    
        #i+=1
    
    for row in range(10):
        for column in range(10):
            if(row < 1 and column < 6):
                die_text_box(row,column+2,column,data,root)
    
    root.mainloop() 
    

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
    width=10
    height=1
    row_distance = row+height*row*3
    column_distance = column+width*column
    die_list = list(data.items())
    die_number = die_Num
    
    
    die = die_list[die_number][0]
    die_value = die_list[die_number][1]
    die_value_top = round(die_value[0],2)
    die_value_bottom = round(die_value[1],2)
    
    display_top_value = tk.Label(root,text=die_value_top,width=width,height=height,borderwidth=1,relief='solid',background='green')
    display_top_value.grid(row=row_distance,column=column_distance)
    
    display_die_number = tk.Label(root,text=die,width=width,height=height,borderwidth=1,relief='solid',background='green')
    display_die_number.grid(row=row_distance+1,column=column_distance)
    
    display_bottom_value = tk.Label(root,text=die_value_bottom,width=width,height=height,borderwidth=1,relief='solid',background='green')
    display_bottom_value.grid(row=row_distance+2,column=column_distance)
    
    print(die,die_value_top,die_value_bottom)
    
    
    
    
def print_data(data):
    print("('die #', [frontside[V], backside[V]])")
    for datas in data.items():
        print(datas)
    
if __name__ == "__main__":
    main()