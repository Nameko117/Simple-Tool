# -*- coding: utf-8 -*-

from pandas import read_excel, concat, DataFrame, merge
from tkinter import Tk, Frame, Label, Button, filedialog, TOP, LEFT
import os

output_columns = ['Customer PO', 'PO LN', 'FAB O/N', 'ON LN', 'Fab Production', 'Customer Production', ' Required Date', ' Confirmed Date', 'Required Qty', 'UoM', 'Hold Flag', 'Order Entry Date', 'Status', 'Micron', 'Order Run Type', ' Fab', 'TTS', 'U/P']

def transform(copr19_path, copr66_path, richtek_path):
    global output_label  
    
    # 讀取copr19文件並轉換
    try:
        copr19 = read_excel(copr19_path)
        result_19 = copr_to_result(copr19)
        result_19['U/P'] = copr19['單   價']
    except:
        print('copr19轉換出錯')
        output_label.configure(text='copr19轉換出錯')
        return
    
    
    # 讀取copr66文件並轉換（可選）
    if copr66_path:
        try:
            copr66 = read_excel(copr66_path)
            result_66 = copr_to_result(copr66)
        except:
            print('copr66轉換出錯')
            output_label.configure(text='copr66轉換出錯')
            return
        result = concat([result_19, result_66], ignore_index=True)
    else:
        result = result_19
        
    # 移除空行並根據單號排序
    result = result.dropna(subset=['Customer PO'])
    result = result.sort_values(['Customer PO'],ascending=True)
    
    try:    
        # 插入richtek資料
        richtek = read_excel(richtek_path)
        richtek = richtek.dropna(subset=['Part No (Customer)'])
        # Hold Flag
        hold_po = richtek[richtek['H']=='C']['P/O No (Customer)'].drop_duplicates()
        for po in hold_po:
            fliter = (result['Customer PO']==po)
            result.loc[fliter, 'Hold Flag'] = 'HOLD'
        # Micron
        CP18_po = DataFrame()
        CP18_po['Customer Production'] = richtek['Part No (Customer)']
        CP18_po['Micron'] = '0.' + richtek['Part No (Key Foundry)'].str[7:9] + 'um'
        CP18_po = CP18_po.drop_duplicates()
        for i in CP18_po.index:
            fliter = (result['Customer Production']==CP18_po['Customer Production'][i])
            result.loc[fliter, 'Micron'] = CP18_po['Micron'][i]
        # Fab
        fab_po = DataFrame()
        fab_po['Customer Production'] = richtek['Part No (Customer)']
        fab_po['Fab'] = richtek['Fab'].str.replace('-', 'AB')
        fab_po = fab_po.drop_duplicates()
        for i in fab_po.index:
            fliter = (result['Customer Production']==fab_po['Customer Production'][i])
            result.loc[fliter, ' Fab'] = fab_po['Fab'][i]
        
        # 輸出copr和richtek的Required Qty誤差
        WF_Qty_po = DataFrame()
        WF_Qty_po['Customer PO'] = richtek['P/O No (Customer)']
        WF_Qty_po['Customer Production'] = richtek['Part No (Customer)']
        WF_Qty_po['WF Qty'] = richtek['WF Qty']
        WF_Qty_po = WF_Qty_po.groupby(['Customer PO', 'Customer Production']).sum()
        error_result = result[['Customer PO', 'Customer Production', 'Required Qty']]
        error_result = error_result.groupby(['Customer PO', 'Customer Production']).sum()
        error_result = merge(error_result, WF_Qty_po, on=['Customer PO', 'Customer Production'])
        error_result['kf'] = error_result['Required Qty']-error_result['WF Qty']
        error_result = error_result[error_result['kf']!=0]
        error_result = error_result.drop(columns=['kf'])
    except:
        print('richtek轉換出錯')
        output_label.configure(text='richtek轉換出錯')
        return
    
    # 輸出MTC_POsummary
    result = result.sort_values(['Customer Production'],ascending=True)
    result['U/P'] = result['U/P'].fillna(method="ffill")
    result_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx")],initialdir = os.getcwd())
    if '.xlsx' not in result_path:
        result_path += '.xlsx'
    error_result.to_excel(result_path[0:-5]+'_error.xlsx')
    result.to_excel(result_path, index=False)
    output_label.configure(text='輸出成功！')

def copr_to_result(copr):
    result = DataFrame(columns=output_columns)
    result['Customer PO'] = copr['客戶單號']
    result['PO LN'] = copr['品    名'] + '.PO'
    result['Customer Production'] = copr['品    名']
    result[' Required Date'] = copr['預交日期']
    result[' Confirmed Date'] = copr['預交日期']
    result['Required Qty'] = copr['訂單數量'] - copr['已交數量']
    result['UoM'] = 'WF'
    result['Order Entry Date'] = copr['訂單日期']
    result['Status'] = 'Confirmed'
    return result


# =============================================================================
# 
# copr19_path = 'Data/COPR1920210831.xlsx'
# copr66_path = 'Data/COPR6620210831.xlsx'
# richtek_path = 'Data/Richtek 20210831_WipStatus.xlsx'
# result_name = 'result'
# transform(copr19_path, copr66_path, richtek_path, result_name)
# 
# =============================================================================

# 建立事件處理函式（event handler），透過元件 command 參數存取
def echo_hello():
    print('hello world :)')
    
def open_file_copr19():
    file_path = filedialog.askopenfilename(title=u'選擇檔案',filetypes=[("Excel files","*.xlsx")],initialdir=os.getcwd())
    global copr19_path_label, copr19, output_label
    if file_path is not None:
        copr19 = file_path
        copr19_path_label.configure(text=file_path)
    else:
        copr19_path_label.configure(text='檔案錯誤')
    output_label.configure(text='')
def open_file_copr66():
    file_path = filedialog.askopenfilename(title=u'選擇檔案',filetypes=[("Excel files","*.xlsx")],initialdir=os.getcwd())
    global copr66_path_label, copr66, output_label
    if file_path is not None:
        copr66 = file_path
        copr66_path_label.configure(text=file_path)
    else:
        copr66_path_label.configure(text='檔案錯誤')
    output_label.configure(text='')
def open_file_richtek():
    file_path = filedialog.askopenfilename(title=u'選擇檔案',filetypes=[("Excel files","*.xlsx")],initialdir=os.getcwd())
    global richtek_path_label, richtek, output_label
    if file_path is not None:
        richtek = file_path
        richtek_path_label.configure(text=file_path)
    else:
        richtek_path_label.configure(text='檔案錯誤')
    output_label.configure(text='')
        
def save_file():
    global copr19, copr66, richtek
    transform(copr19, copr66, richtek)

# 建立主視窗和 Frame（把元件變成群組的容器）
window = Tk()
window.title('小工具')
window.geometry('500x200')

# copr19元件
copr19 = ''
copr19_frame = Frame(window)
copr19_frame.pack(side=TOP)
copr19_label = Label(copr19_frame, text='copr19: ')
copr19_label.pack(side=LEFT)
copr19_button = Button(copr19_frame, text='選擇檔案', command=open_file_copr19)
copr19_button.pack(side=LEFT)
copr19_path_label = Label(window, text='')
copr19_path_label.pack(side=TOP)

# copr66元件
copr66 = ''
copr66_frame = Frame(window)
copr66_frame.pack(side=TOP)
copr66_label = Label(copr66_frame, text='copr66: ')
copr66_label.pack(side=LEFT)
copr66_button = Button(copr66_frame, text='選擇檔案', command=open_file_copr66)
copr66_button.pack(side=LEFT)
copr66_path_label = Label(window, text='')
copr66_path_label.pack(side=TOP)

# richtek元件
richtek = ''
richtek_frame = Frame(window)
richtek_frame.pack(side=TOP)
richtek_label = Label(richtek_frame, text='richtek: ')
richtek_label.pack(side=LEFT)
richtek_button = Button(richtek_frame, text='選擇檔案', command=open_file_richtek)
richtek_button.pack(side=LEFT)
richtek_path_label = Label(window, text='')
richtek_path_label.pack(side=TOP)

# 輸出元件
output_button = Button(window, text='輸出結果', command=save_file)
output_button.pack(side=TOP)
output_label = Label(window, text='')
output_label.pack(side=TOP)

# 運行主程式
window.mainloop()
