from JobHelper import *
import pop3test
import os
from tkinter import *
from tkinter import filedialog
import webbrowser as web
from tkinter import messagebox

#窗口属性
window = Tk()
window.title("收极")
window.geometry('800x580')
try:
    window.iconbitmap(default = r'.\images\tubiao.ico')
except:
    pass
path1 = StringVar()
path2 = StringVar()
counter = 0   #保存函数用，勿删



def do_job():
    global counter
    counter+=1

def selectPath1():
    global path_1
    path_1 = filedialog.askdirectory( )  #获得文件夹地址
    path1.set(path_1)

def selectPath2():
    global path_2
    path_2 = filedialog.askopenfilename( ) #获得文件地址
    path2.set(path_2)

def selectPath5():
    copytoclipboard1(path_1, path_2)
    copy_01.config(text="复制成功")

def selectPath4():
    copytoclipboard2(path_1, path_2)
    copy_02.config(text="复制成功")

def produceinfo():
    try:
        outcome_list = greatoutcome(path_1,path_2)
    except NameError:
        messagebox.showwarning(title='生成结果不成功', message='未选择文件夹')
    aerro = outcome_list[2]#如果excel文件没有关闭，返回 1
    aerro1 = outcome_list[3]
    if aerro == 1:
        messagebox.showwarning(title='生成结果不成功',message='excel文件未关闭')
    elif aerro1 == 1:
        messagebox.showwarning(title='生成结果不成功', message='目标文件夹为空')
    else:
        #结果子窗口
        Child_window = Toplevel(window)
        Child_window.title("未交作业名单")
        Child_window.geometry('600x500')
        # 四、显示未完成的作业
        global var1
        global var2
        global copy_01
        global copy_02
        var1 = StringVar()
        var2 = StringVar()
        a = var1.set(outcome_list[0])
        var2.set(outcome_list[1])
        Label(Child_window, text="未交名单", font=('微软雅黑', 18,)).place(x=110, y=20, width=130, height=50)
        Listbox(Child_window, bd=1, listvariable=var1).place(x=75, y=70, height=350, width=200)
        # 显示无法识别文件
        Label(Child_window, text="不能识别的名单", font=('微软雅黑', 13,)).place(x=365, y=20, width=130, height=50)
        Listbox(Child_window, bd=1, listvariable=var2).place(x=328, y=70, height=350, width=200)

        Button(Child_window, text="复制名单", fg='white',cursor="hand2", bg='#548B54', font=('微软雅黑', 13, 'bold'),
               command=selectPath5).place(x=75, y=430, width=100, height=40)
        copy_01 = Label(Child_window, text="", font=('微软雅黑', 11,))
        copy_01.place(x=195, y=430, width=100, height=50)
        Button(Child_window, text="复制名单", fg='white',cursor="hand2", bg='#548B54', font=('微软雅黑', 13, 'bold'),
               command=selectPath4).place(x=328, y=430, width=100, height=40)  # 复制函数未确定
        copy_02 = Label(Child_window, text="", font=('微软雅黑', 11,))
        copy_02.place(x=447, y=430, width=100, height=50)
        Child_window.resizable(0, 0)  # 防止用户自己调整尺寸
        Child_window.mainloop()


def downlond():
    """创建下载邮件子窗口"""
    Child_window = Toplevel(window)
    Child_window.attributes("-toolwindow", True)#把窗口设置为工具子窗口
    Child_window.attributes("-topmost", True)#让窗口置顶
    #Child_window.attributes('-alpha', 0.8)#改变窗口透明度
    Child_window.title("邮件下载")
    Child_window.geometry('500x400')

    def login_1():
        """登录邮箱服务器"""
        mail_1 = pop3test.mail(var_pop.get(), var_username.get(), var_passward.get())
        flag = mail_1.login()
        print(flag)
        if flag == 0:
            Label(Child_window, text="登录失败").place(x=400, y=110)
        else:
            Label(Child_window, text="登录成功").place(x=400, y=110)

    def selectPath():
        """选择文件存放路径"""
        Child_window.wm_attributes("-topmost", 0)
        path_1 = filedialog.askdirectory()  # 获得文件地址
        var_path.set(path_1)
        Child_window.wm_attributes("-topmost", 1)

    def download_1():
        """下载指定主题的所有邮件附件"""
        try:
            chdir(var_path.get())
        except FileNotFoundError:
            messagebox.showwarning(title='亲~请注意', message='请选择正确的存放文件夹')
            return
        mail_1 = pop3test.mail(var_pop.get(), var_username.get(), var_passward.get())
        server, mailnumber = mail_1.login()
        for i in range(1, mailnumber + 1):
            length = round((i / mailnumber) * 20)
            Label(Child_window, text="下载中    <" + '.' * length + '>').place(x=240, y=260)
            Child_window.update()
            try:
                mail_1.downloadfile(var_subject.get(), var_path.get(), i)
            except FileNotFoundError:
                messagebox.showwarning(title='亲~请注意', message='请选择正确的存放文件夹')
        Label(Child_window, text="下载完毕").place(x=240, y=260)
        server.quit()

    global count_22
    count_22 = 0

    def download_2():
        """逐份下载指定主题的邮件"""
        try:
            chdir(var_path.get())
        except FileNotFoundError:
            messagebox.showwarning(title='亲~请注意', message='请选择正确的存放文件夹')
            return
        global count_22
        mail_1 = pop3test.mail(var_pop.get(), var_username.get(), var_passward.get())
        server, mailnumber = mail_1.login()
        count_21 = mailnumber - count_22
        count_22 += 1
        if count_21 >= 1:
            try:
                mail_1.downloadfile(var_subject.get(), var_path.get(), count_21)
            except FileNotFoundError:
                messagebox.showwarning(title='亲~请注意', message='请选择正确的存放文件夹')
            Label(Child_window, text="已下载最近第" + str(count_22) + "邮件").place(x=240, y=300)
        else:
            Label(Child_window, text="已下载最近所有邮件").place(x=240, y=300)
            server.quit()

    def reloadcount_22():
        global count_22
        count_22 = 0

    # 如何给自己的邮箱开启pop3服务
    def openpop3manual():
        web.open_new_tab("https://jingyan.baidu.com/article/c85b7a64be9284003bac9535.html")

    def openconnect():
        web.open_new_tab("http://47.111.70.60/")

    def openknowledge():
        web.open_new_tab("http://47.111.70.60/")

    # 菜单栏
    menubar = Menu(Child_window)  # 初始化菜单栏
    functionmenu = Menu(menubar, tearoff=0)  # 定义一个父菜单
    menubar.add_cascade(label='帮助', menu=functionmenu)  # 在菜单栏上的父菜单命名
    functionmenu.add_command(label='开启pop3服务指导', command=openpop3manual)
    functionmenu.add_command(label='联系我们', command=openconnect)
    functionmenu.add_command(label='学生论坛', command=openknowledge)
    Child_window.config(menu=menubar)

    Label(Child_window, text='>>>邮件下载工具(如何打开pop3服务请在帮助中获取)<<<', fg='black', bg='white',
          height=1, width=50).pack(side='top')
    # pop3服务器
    var_pop = StringVar()
    var_pop.set('pop.qq.com')
    Label(Child_window, text='pop3服务器', fg='white', bg='#104E8B', height=1, width=12).place(x=20, y=30, anchor='nw')
    e1 = Entry(Child_window, textvariable=var_pop)
    e1.place(x=130, y=30, anchor='nw')
    # 用户名
    var_username = StringVar()
    var_username.set('1825819420@qq.com')
    Label(Child_window, text='邮箱账号', fg='white', bg='#104E8B', height=1, width=12).place(x=20, y=70, anchor='nw')
    e2 = Entry(Child_window, textvariable=var_username)
    e2.place(x=130, y=70, anchor='nw')
    # 密码
    var_passward = StringVar()
    var_passward.set("oosxvjkfuyoiegdh")
    Label(Child_window, text='邮箱密码',fg='white', bg='#104E8B', height=1, width=12).place(x=20, y=110, anchor='nw')
    e3 = Entry(Child_window, textvariable=var_passward, show="*")
    e3.place(x=130, y=110, anchor='nw')
    # 登录按钮
    Button(Child_window, text="登录邮箱",bg="#B22222",fg='white',cursor="hand2", relief='groove', command=login_1,
           height=1, width=8).place(x=300, y=100)

    # 邮件主题
    var_subject = StringVar()
    var_subject.set("数学作业")
    Label(Child_window, text='邮件主题', fg='white', bg='#104E8B', height=1, width=12).place(x=20, y=180, anchor='nw')
    e3 = Entry(Child_window, textvariable=var_subject)
    e3.place(x=130, y=180, anchor='nw')
    # 附件存放路径
    var_path = StringVar()
    var_path.set("请选择文件存放路径")
    Label(Child_window, text='文件存放路径', fg='white', bg='#104E8B', height=1, width=12).place(x=20, y=210, anchor='nw')
    e3 = Entry(Child_window, textvariable=var_path)
    e3.place(x=130, y=210, anchor='nw')
    Button(Child_window, text='选择路径', bg="#B22222",fg='white',cursor="hand2", relief='groove',
           command=selectPath, height=1, width=8).place(x=300, y=200)

    # 下载邮件
    Button(Child_window, text="下载对应主题邮件的附件", bg="#B22222",fg='white',cursor="hand2", relief='groove',
           command=download_1, height=1, width=30).place(x=10, y=260)
    Button(Child_window, text="下载最近一封邮件的附件(不限制主题)", bg="#B22222",fg='white',cursor="hand2", relief='groove',
           command=download_2, height=1, width=30).place(x=10, y=297)
    Button(Child_window, text="重置",bg="#B22222",fg='white',cursor="hand2", relief='groove',
           command=reloadcount_22, height=1, width=10).place(x=357, y=297)
    Child_window.resizable(0, 0)  # 防止用户自己调整尺寸
    Child_window.mainloop()

def Rename():
    """
    重命名文件夹里的所有文件
    """
    try:
        renamef(path_1, path_2)
        os.startfile(path_1)
    except NameError:
        messagebox.showwarning(title='生成结果不成功', message='未选择文件夹')


def open_right_01():
    l_B1.config(fg='white',bg='#CD2626')
    l_B2.config(fg='black', bg='white')
    Label(right_frm, text="交作业", bg="white", font=('微软雅黑', 30,)).place(x=270, y=30, width=130, height=50)
    Label(right_frm, text="作业路径:", bg="white", font=('微软雅黑', 16,)).place(x=270, y=230, width=90, height=50)
    Entry(right_frm, textvariable=path1, bg="#FAFAFA", width=22, insertbackground='blue').place(x=380, y=245)
    Button(right_frm, text="路径选择", bg="white", command=selectPath1,cursor="hand2", relief='groove', font=('微软雅黑', 13)).place(x=570,y=235,
                                                                                                              width=90,
                                                                                                              height=40)
    # 二、excel文件选择模块
    Label(right_frm, text="班级excel文件:", bg="white", font=('微软雅黑', 16,)).place(x=230, y=300, width=140, height=50)
    Entry(right_frm, textvariable=path2, bg='#FAFAFA', width=22, insertbackground='blue').place(x=380, y=315)
    Button(right_frm, text="路径选择", bg="white",width=20, height=6,cursor="hand2", command=selectPath2, relief='groove', font=('微软雅黑', 13,)).place(
        x=570, y=305, width=90, height=40)
    Button(right_frm, text="看哪个小可爱没有交作业", command=produceinfo,cursor="hand2", relief='flat', fg='white', bg='#104E8B',
           font=('微软雅黑', 15, 'bold')).place(x=380, y=400)
def open_right_02():
    l_B1.config(fg='black',bg='white')
    l_B2.config(fg='white', bg='#104E8B')
    Label(right_frm01, text="重命名", bg="white", font=('微软雅黑', 30,)).place(x=270, y=30, width=130, height=50)
    Label(right_frm01, text="作业路径:", bg="white", font=('微软雅黑', 16,)).place(x=270, y=230, width=90, height=50)
    Entry(right_frm01, textvariable=path1, bg="#FAFAFA", width=22, insertbackground='blue').place(x=380, y=245)
    Button(right_frm01, text="路径选择", bg="white",cursor="hand2", command=selectPath1, relief='groove', font=('微软雅黑', 13)).place(x=570, y=235, width=90, height=40)
    # 二、excel文件选择模块
    Label(right_frm01, text="班级excel文件:", bg="white", font=('微软雅黑', 16,)).place(x=230, y=300, width=140, height=50)
    Entry(right_frm01, textvariable=path2, bg='#FAFAFA', width=22, insertbackground='blue').place(x=380, y=315)
    Button(right_frm01, text="路径选择", bg="white", width=20, height=6, command=selectPath2,cursor="hand2", relief='groove',font=('微软雅黑', 13,)).place(x=570, y=305, width=90, height=40)
    Button(right_frm01, text="帮没命名好的小可爱命名", command=Rename, relief='flat',cursor="hand2", fg='white',bg='#CD2626',font=('微软雅黑', 15, 'bold')).place(x=380, y=400)

def open_Official_website():    #进入官网
    web.open_new_tab("http://47.111.70.60/")
def Getting_Started():      #进入入门指南
    web.open_new_tab("http://47.111.70.60/manual.html")
def about_us():             #关于我们
    web.open_new_tab("http://47.111.70.60/about%20us.html")
def Check_apdate():
    messagebox.showinfo(title='检查更新', message='当前版本为最新版')
#1、菜单栏
menubar = Menu(window)       #初始化菜单栏
functionmenu = Menu(menubar,tearoff=0)        #定义一个父菜单
menubar.add_cascade(label='功能(f)',menu=functionmenu) #在菜单栏上的父菜单命名
functionmenu.add_command(label='下载邮件',command=downlond)
functionmenu.add_command(label='常用通知模板',command=do_job)
functionmenu.add_command(label='学生论坛',command=do_job)
#functionmenu.add_separator()   #加一条线用于分离
#另一个菜单栏
helpmenu = Menu(menubar,tearoff=0)        #定义一个父菜单
#menubar=菜单栏,tearoff 分开与不分开的一个区别
menubar.add_cascade(label='帮助(h)',menu=helpmenu) #在菜单栏上的父菜单命名
helpmenu.add_command(label='进入官网',command=open_Official_website)
helpmenu.add_command(label='入门指南',command=Getting_Started)
helpmenu.add_command(label='检查更新',command=Check_apdate)
helpmenu.add_command(label='关于我们',command=about_us)
#helpmenu.add_separator()  #加一条线用于分离

window.config(menu=menubar)   #保存Menu一直成立

#2、搭框架
main_frm = Frame(window).place()
left_frm = Frame(main_frm,background='#C1C1C1',width=200,height=650).place(x=0,y=0)
right_frm = Frame(main_frm,background='white',width=600,height=650).place(x=200,y=0)
right_frm01 = Frame(main_frm,background='white',width=600,height=650).place(x=200,y=0)



#3、左框架的内容
try:
    canvas_1 = Canvas(window, bg='#C1C1C1', relief='flat', height=176, width=200)  # 应为logo的大小（待修改）
    image_file = PhotoImage(file=r'.\images\tubiao.gif')
    canvas_1.create_image(0, 0, anchor='nw', image=image_file)  # 待填图片，未修改
    canvas_1.create_rectangle(0, 0, 200, 176, outline='#C1C1C1', width=4)
    canvas_1.place(x=-3, y=-3)
except:
    pass
#Label(right_frm,text = "功能选择↓（按钮）",bg="white",font=('微软雅黑',13,),fg='black').place(x=50,y=170,width=150,height=40)
global l_B1
l_B1 = Button(left_frm,text='收作业/材料',bg="white",font=('微软雅黑',16,),fg='black',relief='ridge',cursor="hand2",command=open_right_01)
l_B1.place(x=50,y=228,width=150,height=40)
global l_B2
l_B2 = Button(left_frm,text='统一格式重命名',bg="white",font=('微软雅黑',14),relief='ridge',cursor="hand2",command=open_right_02)
l_B2.place(x=50,y=295,width=150,height=40)
#canvas_2 = Canvas(window,bg='blue',height=350,width=200)   #待填图片，未修改
#image_file_1 = PhotoImage(file=r'C:\Users\打开新世界的大门\Desktop\测试文件夹\18_T_5_J_9_ARM21_3.gif')
#canvas_2.create_image(0,0,anchor='nw',image=image_file_1)
#canvas_2.place(x=0,y=280)
Label(left_frm,text = "QQ群：786418342",background='#C1C1C1').place(x=20,y=490)
Label(left_frm,text = "邮箱：3169162808@qq.com",background='#C1C1C1').place(x=20,y=515)
Label(left_frm,text = "有问题请及时联系我们",background='#C1C1C1').place(x=20,y=540)

window.resizable(0,0)#防止用户自己调整尺寸
window.mainloop()