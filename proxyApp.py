'''mporting Libraries'''

import pandas as pd
import os
import tkinter
from tkinter import filedialog

'''Initializing Global Variables'''

flag = True
minClassHrs = 60
minAttendence = 75
folder_path = None
file_path = None
save_path = None
error = None

'''Main Compiling Function'''

def main(df):
    col_n, col_a, col_t = df.columns  # just for readability: name, action, time
    i_n, i_a, i_t = range(3)
    df[col_t] = pd.to_datetime(df[col_t])  # for calculating
    meeting_end, meeting_begin = max(df[col_t]), min(df[col_t])
    meeting_duration_secs = (meeting_end - meeting_begin).total_seconds()
    names = sorted(set(df[col_n]))
    attendance = {}  # summarize time deltas per name
    for name in names:
        arr = df[df[col_n] == name].values  # ndarray of current attendee slice
        assert arr[0][i_a].startswith("Joined")
        attendance[name] = 0.
        for i in range(len(arr) - 1):
            row_1, row_2 = arr[i], arr[i+1]
            if row_1[i_a].startswith("Joined") and row_2[i_a] == "Left":
                attended = row_2[i_t] - row_1[i_t]
                attendance[name] += attended.total_seconds()
        if arr[-1][i_a] != "Left":
            attended = meeting_end - arr[-1][i_t]
            attendance[name] += attended.total_seconds()

    
    return(attendance,meeting_duration_secs)

'''Function for getting suitable output'''

def reCast(tdf):
    global minAttendence
    #tdf.fillna('Absent',inplace=True)
    tdf.iloc[:,2:] = tdf.iloc[:,2:].replace({'Present': 1 ,'Absent': 0})
    total = len(tdf.columns)-2
    tdf.iloc[:,2:] = tdf.iloc[:,2:].astype(int)
    tdf['Class Attended'] = tdf.iloc[:,2:].sum(axis = 1)
    tdf['Class Attended'] = tdf['Class Attended'].astype(int)
    for index,d in tdf.iterrows():
        if d['Class Attended']>=total*minAttendence/100:
            tdf.loc[index,'Remark'] = 'Safe'
        else:
            tdf.loc[index,'Remark'] = 'Detained'
    tdf.iloc[:,2:-2] = tdf.iloc[:,2:-2].replace({1: 'Present' ,0:'Absent'})
    for index , d in tdf.iterrows():
        tdf.loc[index,'Class Attended'] = str(tdf.loc[index,'Class Attended'])+'/'+str(total)
    tdf.rename(columns = {'Class Attended':'Class Attended/Total'})
    
    return(tdf)

''' Output Formatting '''

'''
def highlight(val):
     
    if val == 'Present':
        cr = 'blue'
    elif val == 'Absent':
        cr = 'orange'
    elif val == 'Safe':
        cr = 'green'
    elif val == 'Detained':
        cr = 'red'
    else:
        cr = 'black'
        
    return 'color : %s' % cr
'''

'''Mater Function'''
def masterFunction():
    global flag,error
    if flag == True:
       
        myList = os.listdir(folder_path)
        
        file_df = []
        if not myList:
            flag = False
            error = 'Selected folder is Empty..!!'
        else:
            for file in myList:
                try:
                    temp_df = pd.read_csv(f'{folder_path}/{file}',sep='\t',encoding='utf-16')
                except:
                    pass
                else:
                    date = temp_df['Timestamp'][0].split(',')[0]
                    temp_df['Full Name'] = temp_df['Full Name'].str.replace('Mr. ','')
                    temp_df['Full Name'] = temp_df['Full Name'].str.replace('Ms. ','')
                    temp_df['Full Name'] = temp_df['Full Name'].str.upper()
                    temp_df = main(temp_df),date
                    file_df.append(temp_df)
            
            if not file_df:
                flag = False
                error = 'No suitable File format in folder..!!'
            else:
                #Main Sheet
                try:
                    sheet = pd.read_excel(file_path,sheet_name = None,engine='openpyxl')
                except:
                    flag = False
                    error = 'Selected Excel Worksheet is Empty..!!'
                else:
                    for key in sheet:
                        mdf = sheet[key]
                        mdf['Student Name'] = mdf['Student Name'].str.replace('Mr. ','')
                        mdf['Student Name'] = mdf['Student Name'].str.replace('Ms. ','')
                            
                        for name in mdf['Student Name']:
                            for (dtf,check),date in file_df:
                                if name in dtf.keys():
                                    if(dtf[name]>=check*minClassHrs/100):
                                        mdf.set_index('Student Name',inplace = True)
                                        mdf.loc[name,date] = 'Present'
                                        mdf.reset_index(inplace = True)
                                    else:
                                        mdf.set_index('Student Name',inplace = True)
                                        mdf.loc[name,date] = 'Absent'
                                        mdf.reset_index(inplace = True)
                                else:
                                    mdf.set_index('Student Name',inplace = True)
                                    mdf.loc[name,date] = 'Absent'
                                    mdf.reset_index(inplace = True)    
                            
                        sheet[key] = reCast(sheet[key])
                        #sheet[key] = sheet[key].style.applymap(highlight)
                        sheet[key].to_excel(f'{save_path}/{key}Attendence.xlsx',index = False)
                        #print(f'Successfully Saved...{key}Attendence.xlsx')
                        
                    #print('Successfully Completed..!!')

'''Tkinter GUI Functions'''

def getInputs():
    global minClassHrs
    global minAttendence
    try:
        minClassHrs = int(entry1.get())
        minAttendence = int(entry2.get())
    except:
        pass
    else:
        root.destroy()

def open_folder():
    global folder_path 
    folder_path = filedialog.askdirectory()
    if folder_path != '':
        cfm3 = tkinter.Label(root,text=u'\u2713',fg='green',bg='#ffffff')
        cfm3.place(x=275,y=150)
    
def open_file():
    global file_path 
    file_path = filedialog.askopenfilename()
    if file_path != '':
        cfm1 = tkinter.Label(root,text=u'\u2713',fg='green',bg='#ffffff')
        cfm1.place(x=275,y=190)

def save_file():
    global save_path
    save_path = filedialog.askdirectory()
    if save_path != '':
        cfm2 = tkinter.Label(root,text=u'\u2713',fg='green',bg='#ffffff')
        cfm2.place(x=275,y=230)
    
def exitBtn():
    global flag
    flag = False
    root.destroy()

'''Tkinter GUI Interface'''

while (folder_path == None or file_path == None or save_path == None ) and flag==True :    
    root = tkinter.Tk()

    #root.iconbitmap('LOGO.ico')
    #img = tkinter.PhotoImage(file = 'logo.png')
    #photo = tkinter.Label(root, image = img)
    headLabel = tkinter.Label(root,text = 'Proxy App',fg = 'Black', font = ('Forte',18))
    label1 = tkinter.Label(root,text = 'Minimum Class Hours (%) : ',font = ('Tw Cen MT Condensed',12))
    label2 = tkinter.Label(root,text = 'Minimum Attendence Required (%) : ',font = ('Tw Cen MT Condensed',12))
    label3 = tkinter.Label(root,text = u'\u00A9' + '  Pratyush Lokhande', font = ('Amazed Breath',15))
    entry1 = tkinter.Entry(root,bd = 5)
    entry2 = tkinter.Entry(root,bd = 5)
    button1 = tkinter.Button(root,text='Select Attendence Files Folder',width = 30,bg='#ffffff',font = ('Tw Cen MT Condensed',12),activebackground='#00ff00',command=open_folder)
    button2 = tkinter.Button(root,text='Select Student List',width = 30,bg='#ffffff',font = ('Tw Cen MT Condensed',12),activebackground='#00ff00',command=open_file)
    button4 = tkinter.Button(root,text='Select Save Path',width = 30,bg='#ffffff',font = ('Tw Cen MT Condensed',12),activebackground='#00ff00',command= save_file)
    button3 = tkinter.Button(root,text='START COMPILING',width = 30,bg='#ffffff',fg='green',font = ('Tw Cen MT Condensed',12),activebackground='#00ff00',command=getInputs)
    '''
    style = ttk.Style()
    style.theme_use('default')
    style.configure("grey.Horizontal.TProgressbar",background='green')
    bar = Progressbar(root,length=220,style='grey.Horizontal.TProgressbar')
    bar['value'] = progress
    '''
    button5 = tkinter.Button(root,text='EXIT',width = 30,bg='#ffffff',fg='red',font = ('Tw Cen MT Condensed',12),activebackground='#00ff00',command=exitBtn)
    headLabel.place(x=120,y=10)
    entry1.place(x=210,y=60)
    entry2.place(x=210,y=100)
    button1.place(relx=0.5,y=160, anchor = 'center')
    button2.place(relx=0.5,y=200, anchor = 'center')
    button4.place(relx=0.5,y=240, anchor = 'center')
    button3.place(relx=0.5,y=280, anchor = 'center')
    label1.place(x=10,y=65)
    label2.place(x=10,y=105)
    #bar.place(x=70,y=330)
    button5.place(relx=0.5,y=330, anchor = 'center')
    label3.place(x=130,y=380)
    #photo.place(x=150,y=50)
    root.geometry('350x420+120+120')
    root.title('Proxy')
    root.protocol('WM_DELETE_WINDOW', exitBtn)
    root.mainloop()
    

'''Calling Master Function'''
masterFunction()



'''Final Dialog Window to display success message'''
root2 = tkinter.Tk()
root2.title('Result')
if flag==True:
    flabel = tkinter.Label(text=u'\u2713'+'  Successfully Completed..!!',fg = 'green',font = ('Tw Cen MT Condensed',15))
    flabel.place(relx=0.5,rely=0.4, anchor = 'center')
elif error==None:
    flabel = tkinter.Label(text=u'\u274C'+'  Action Interrupted..!!',fg = 'red',font = ('Tw Cen MT Condensed',15))
    flabel.place(relx=0.5,rely=0.4, anchor = 'center')
else:
    flabel = tkinter.Label(text = u'\u274C'+" "+error,fg = 'blue',font = ('Tw Cen MT Condensed',15))
    flabel.place(relx=0.5,rely=0.4, anchor = 'center')
labelc = tkinter.Label(root2,text = u'\u00A9' + '  Pratyush Lokhande', font = ('Amazed Breath',12))
labelc.place(x=90,y=70)
root2.geometry('250x100+100+100')
root2.mainloop()
