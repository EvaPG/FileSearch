import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import os
import shutil
import threading
import ctypes
import inspect
import time
from funcSearchFileContent import txtContentFindString
from funcSearchFileContent import wordDocContentFindString
from funcSearchFileContent import wordDocxContentFindString
from funcSearchFileContent import excelContentFindString

temp_path = os.path.join(os.path.expanduser("~"), 'Desktop')+"\\SearchTemp"
intSearchMatch=0
intIsAllFiles=0
commonFileTypes = ['txt', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'vsd', 'vsdx', 'jpg', 'jpge', 'png', 'bmp',
                   'gif', 'mp3', 'wma', 'wav', 'ape', 'real', 'rm', 'avi', 'mp4', 'mkv', 'wmv', 'rmvb', 'qlv', 'mov',
                   'flv', 'rar', 'zip', '7z']
canMactchContentFileTypes = ['.txt', '.doc', '.docx', '.xls', '.xlsx']
checkbuttonFiltTypeObj=[]
checkbuttonFiltTypeVar=[]
searchFileTypes=[]
Thread=[]
strExportSearchResultPath=''

def selectSearchPath():
    strSearchPath.set(filedialog.askdirectory())
    # print(strSearchPath.get().lstrip())
def selectSearchMatch():
    global intSearchMatch
    if (intMatchFileName.get() == 1) & (intMatchFileContent.get() == 1):
        intSearchMatch=3#匹配文件名和文件内容
    elif (intMatchFileName.get() == 1) & (intMatchFileContent.get() == 0):
        intSearchMatch=2#匹配文件名
    elif (intMatchFileName.get() == 0) & (intMatchFileContent.get() == 1):
        intSearchMatch=1#匹配文件内容
    else:
        intSearchMatch=0#全不匹配，此时无法搜索

def selectAllFiles():
    global intIsAllFiles
    if(intAllFiles.get()==1):
        intIsAllFiles=1
        checkbuttonAllTypes.config(state='disabled')
        for obj in checkbuttonFiltTypeObj:
            obj.config(state='disabled')
    else:
        intIsAllFiles=0
        checkbuttonAllTypes.config(state='normal')
        for obj in checkbuttonFiltTypeObj:
            obj.config(state='normal')

def selectAllTypes():
    if(intAllTypes.get()==1):
        for obj in checkbuttonFiltTypeObj:
            obj.select()
    else:
        for obj in checkbuttonFiltTypeObj:
            obj.deselect()

def _initFileTypePanel():
    for i in range(len(commonFileTypes)):
        checkbuttonFiltTypeVar.insert(i,tk.IntVar())
        checkbuttonFiltTypeObj.insert(i,tk.Checkbutton(frameFileType, text=commonFileTypes[i],variable=checkbuttonFiltTypeVar[i], onvalue=1, offvalue=0))
        checkbuttonFiltTypeObj[i].grid(row=int(i / 11) + 3, column=i % 11, sticky='w')

def startSearch():
    if(trim(strSearchPath.get().lstrip())==''):
        messagebox.askokcancel('提醒', "请选择搜索路径！")
        return
    if (trim(strSearchContent.get().lstrip()) == ''):
        messagebox.askokcancel('提醒', "搜索内容不能为空！")
        return
    if (intSearchMatch == 0):
        messagebox.askokcancel('提醒', "请选择搜索匹配规则！")
        return

    del searchFileTypes[:]
    for i in range(len(checkbuttonFiltTypeVar)):
        if(checkbuttonFiltTypeVar[i].get()==1):
            searchFileTypes.append('.'+commonFileTypes[i])
    if(intIsAllFiles ==0 and len(searchFileTypes)==0):
        messagebox.askokcancel('提醒', "请选择文件类型！")
        return

    buttonStartSearch.config(state='disabled')
    buttonStopSearch.config(state='normal')
    buttonExportSearchResult.config(state='disabled')
    oldSearchResult = treeviewSearchResult.get_children()
    for row in oldSearchResult:
        treeviewSearchResult.delete(row)
    create_thread(searchMain,'')


def stopSearch():
    if len(Thread)>0:
        for t in Thread:
            stop_thread(t)
    buttonStartSearch.config(state='normal')
    if os.path.exists(temp_path):
        shutil.rmtree(temp_path)

def exportSearchResult():
    create_thread(copyTreeViewListFile, '')

def copyTreeViewListFile():
    strExportSearchResultPath = filedialog.askdirectory()
    if trim(strExportSearchResultPath)!='':
        buttonExportSearchResult.config(state='disabled')
        for rowItem in treeviewSearchResult.get_children():
            if not os.path.exists(os.path.join(strExportSearchResultPath,os.path.split(treeviewSearchResult.item(rowItem, "values")[1])[1])):
                shutil.copyfile(treeviewSearchResult.item(rowItem, "values")[1], os.path.join(strExportSearchResultPath,os.path.split(treeviewSearchResult.item(rowItem, "values")[1])[1]))
        buttonExportSearchResult.config(state='normal')

def searchMain():
    for root, dirs, files in os.walk(strSearchPath.get().lstrip()):
        if(root==temp_path):
            continue
        for file in files:
            filename, type = fileAttr(file)
            if intIsAllFiles==1:
                if intSearchMatch == 1:
                    findByFileContent(root,file,filename,type)
                if intSearchMatch == 2:
                    findByFileName(root,file,filename)
                if intSearchMatch == 3:
                    findByFileNameAndContent(root,file,filename)
            else:
                if type in searchFileTypes:
                    if intSearchMatch==1:
                        findByFileContent(root,file,filename,type)
                    if intSearchMatch==2:
                        findByFileName(root,file,filename)
                    if intSearchMatch==3:
                        findByFileNameAndContent(root,file,filename)
    if os.path.exists(temp_path):
        shutil.rmtree(temp_path)
    buttonStartSearch.config(state='normal')
    buttonStopSearch.config(state='disabled')
    buttonExportSearchResult.config(state='normal')


def findByFileName(root,file,filename):
    if strSearchContent.get().lstrip() in filename:
        treeviewSearchResult.insert('', 'end', value=(len(treeviewSearchResult.get_children())+1, os.path.join(root, file)))

def findByFileContent(root,file,filename,type):
    matchResult = False
    if type in canMactchContentFileTypes:
        if type == ".txt":
            matchResult=txtContentFindString(os.path.join(root, file), strSearchContent.get().lstrip())
        if type == ".doc":
            matchResult=wordDocContentFindString(os.path.join(root, file), filename, strSearchContent.get().lstrip(),temp_path)
        if type == ".docx":
            matchResult=wordDocxContentFindString(os.path.join(root, file), strSearchContent.get().lstrip())
        if type == ".xls" or type == ".xlsx":
            matchResult=excelContentFindString(os.path.join(root, file), strSearchContent.get().lstrip())
    if matchResult:
        treeviewSearchResult.insert('', 'end', value=(len(treeviewSearchResult.get_children()) + 1, os.path.join(root, file)))

def findByFileNameAndContent(root,file,filename):
    matchResult = False
    if strSearchContent.get().lstrip() in filename:
        matchResult=True
    else:
        if type in canMactchContentFileTypes:
            if type == ".txt":
                matchResult = txtContentFindString(os.path.join(root, file), strSearchContent.get().lstrip())
            if type == ".doc":
                matchResult = wordDocContentFindString(os.path.join(root, file), filename, strSearchContent.get().lstrip(),temp_path)
            if type == ".docx":
                matchResult = wordDocxContentFindString(os.path.join(root, file), strSearchContent.get().lstrip())
            if type == ".xls" or type == ".xlsx":
                matchResult = excelContentFindString(os.path.join(root, file), strSearchContent.get().lstrip())
    if matchResult:
        treeviewSearchResult.insert('', 'end', value=(len(treeviewSearchResult.get_children()) + 1, os.path.join(root, file)))

def trim(s):
    if s[:1] != ' ' and s[-1:] != ' ':
        return s
    elif s[:1] == ' ':
        return trim(s[1:])
    else:
        return trim(s[:-1])

def fileAttr(file):
    filename, type = os.path.splitext(file)
    return filename, type

def _async_raise(tid, exctype):
    tid = ctypes.c_long(tid)
    if not inspect.isclass(exctype):
        exctype = type(exctype)
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
    if res == 0:
        raise ValueError("invalid thread id")
    elif res != 1:
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
        raise SystemError("PyThreadState_SetAsyncExc failed")

def create_thread(func,param):
    del Thread[:]
    threadSearc = threading.Thread(target=func,args=(param))
    threadSearc.setDaemon(True)
    Thread.append(threadSearc)
    threadSearc.start()

def stop_thread(thread):
    _async_raise(thread.ident, SystemExit)

windows = tk.Tk()
windows.title('文件搜索工具测试版Version 1.1 By:HeJunjie 2018/08/24')
windows.resizable(width=False,height=False)
screenWidth = windows.winfo_screenwidth()
screenHeight = windows.winfo_screenheight()
windowsWidth = 610
windowsHeight = 580
x = (screenWidth-windowsWidth) / 2
y = (screenHeight-windowsHeight) / 2
windows.geometry("%dx%d+%d+%d" %(windowsWidth,windowsHeight,x,y))

frameBasic = tk.Frame(windows)
tk.Label(frameBasic,text='搜索路径:').grid(row=0,column=0)
strSearchPath = tk.StringVar()
tk.Entry(frameBasic,textvariable=strSearchPath,width=50).grid(row=0,column=1,padx=5)
tk.Button(frameBasic,text='选择搜索路径',command=selectSearchPath).grid(row=0,column=2,columnspan=2,sticky='w')

tk.Label(frameBasic,text='搜索内容:').grid(row=1,column=0)
strSearchContent = tk.StringVar()
tk.Entry(frameBasic,textvariable=strSearchContent,width=50).grid(row=1,column=1,padx=5)
intMatchFileName=tk.IntVar()
tk.Checkbutton(frameBasic,text='文件名',variable=intMatchFileName,onvalue=1,offvalue=0,command=selectSearchMatch).grid(row=1,column=2)
intMatchFileContent=tk.IntVar()
tk.Checkbutton(frameBasic,text='文件内容',variable=intMatchFileContent,onvalue=1,offvalue=0,command=selectSearchMatch).grid(row=1,column=3)
tk.Label(frameBasic,text='(文件内容搜索目前只支持txt,doc,docx,xls,xlsx)').grid(row=2,column=0,columnspan=4,sticky='e')
frameBasic.grid(row=0,column=0,padx=10,pady=10)


frameFileType = tk.LabelFrame(windows,text='文件类型')
tk.Label(frameFileType,text='所有类型:').grid(row=0,column=0,sticky='e')
intAllFiles = tk.IntVar()
checkbuttonAllFiles = tk.Checkbutton(frameFileType, text='(本机所有文件)', variable=intAllFiles,onvalue=1, offvalue=0,command=selectAllFiles)
checkbuttonAllFiles.grid(row=0,column=1,columnspan=10,sticky='w')

tk.Label(frameFileType,text='指定类型:').grid(row=2,column=0)
intAllTypes = tk.IntVar()
checkbuttonAllTypes=tk.Checkbutton(frameFileType, text='全选', variable=intAllTypes,onvalue=1, offvalue=0,command=selectAllTypes)
checkbuttonAllTypes.grid(row=2,column=1,columnspan=10,sticky='w')
_initFileTypePanel()

frameFileType.grid(row=1,column=0,padx=10)

frameSearchResult = tk.LabelFrame(windows,text='搜索结果')
clframeSearchResult = tk.Frame(frameSearchResult)
crframeSearchResult = tk.Frame(frameSearchResult)

xscrollbarSearchResult = tk.Scrollbar(clframeSearchResult,orient='horizontal')
yscrollbarSearchResult = tk.Scrollbar(crframeSearchResult)

treeviewSearchResult = ttk.Treeview(clframeSearchResult, yscrollcommand=yscrollbarSearchResult.set,xscrollcommand=xscrollbarSearchResult.set,show="headings",columns=('No','File'))
treeviewSearchResult.column('No',width=50)
treeviewSearchResult.column('File',width=518)
treeviewSearchResult.heading('No',text='编号')
treeviewSearchResult.heading('File',text='文件')

treeviewSearchResult.pack( side='top',fill='both')
xscrollbarSearchResult.pack(side='bottom', fill='x')
xscrollbarSearchResult.config(command=treeviewSearchResult.xview)

yscrollbarSearchResult.pack(side='right',fill='y',expand='yes')
yscrollbarSearchResult.config(command=treeviewSearchResult.yview)

clframeSearchResult.pack(side='left')
crframeSearchResult.pack(side='right',fill='y')

frameSearchResult.grid(row=2,column=0,padx=10,pady=10)

frameButton = tk.Frame(windows)
buttonStartSearch=tk.Button(frameButton,text='开始搜索',command=startSearch)
buttonStartSearch.grid(row=0,column=0)
buttonStopSearch=tk.Button(frameButton,text='停止搜索',command=stopSearch,state='disabled')
buttonStopSearch.grid(row=0,column=1,padx=10)
buttonExportSearchResult=tk.Button(frameButton,text='复制导出',command=exportSearchResult,state='disabled')
buttonExportSearchResult.grid(row=0,column=3)
frameButton.grid(row=3,column=0)
windows.mainloop()
