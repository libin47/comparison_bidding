import os.path
import tkinter as tk
import tkinter.filedialog as tkf
import tkinter.ttk as ttk
import threading as mt
import sv_ttk
import os
from compfile import compare


class Readme(tk.Frame):
    def __init__(self, root=None):
        super().__init__(root)
        self.root = root
        self.root.title('说明')
        self.frame = ttk.Frame(self.root)
        self.frame.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S), padx=50, pady=50)
        text = "【使用说明】：\n 选择招标文件(最多3份)——>选择投标文件(最多5份)——>选择投标文件对应的标记颜色(点击C按钮)——>点击“运行”以开始。" \
               "\n【参数说明】:\n最小检测长度：对小于此参数长度的子句将不进行重复检测\n" \
               "分割词：处理过程中将以此对句子进行分割成多个子句，再与其他文档进行比较。以正则表达式语法，多个可能的字符以'|'连接\n" \
               "颜色说明：假如投标文件1颜色为红色，投标文件2颜色为绿色，投标文件3颜色为蓝色，\n" \
               "        那么投标文件1输出中与2重复的标记成绿色，与3重复的标成蓝色，其他同理" \
               "\n【额外说明】：\n1、只能接受docx格式的文件，如果源文件为doc或其他格式，需要另存为docx格式才行，直接在文件名中改后缀可能会引发未知错误\n" \
               "2、运行起来没有停止按钮，只能等运行结束后再次点击运行重新运行，参数请谨慎设置\n" \
               "3、运行完毕后会在每个投标文件所在目录下生成两个文件：\n" \
               "    _输出.docx文件是已经将所有重复文本标记上对应颜色的文件\n" \
               "    _重复子句.txt中是检测到与其他投标文件重复且在招标文件中不存在的子句。\n" \
               "4、因为读取本程序读取docx文件的方法原因，子句划分时依次以段落、runs、指定的切割词来划分，\n" \
               "    如果一个子句有重复，会将整个run都标记成对应颜色；\n" \
               "    同时也可能会出现同一句话因为字体样式不同(分属不同run)而在中间切开并标记不同颜色的情况，" \
               "    属于正常情况，大体上应该……不会影响结果。\n" \
               "5、如果某一句话与其他多个投标文件重复，则只会标记成其中一种的颜色。一般是位置靠后的文档对应的颜色。\n" \
               "6、在程序开始执行时请在office中关闭'*_输出.docx'文件，否则会导致无法写入。\n" \
               "7、留心日志输出栏及cmd窗口的错误提示，如果出现其他错误且与作者联系。"
        ttk.Label(self.frame, text=text).grid()



class ColorSet(tk.Frame):
    def __init__(self, app, num, root=None, r=255, g=0, b=0):
        super().__init__(root)
        self.root = root
        self.root.title('颜色选择')
        self.app = app
        self.num = num
        self.r = 255
        self.g = 0
        self.b = 0
        self.frame_color = ttk.Frame(self.root)
        self.frame_color.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S), padx=50, pady=50)
        self.root.configure(background="#%02x%02x%02x" % (int(self.r), int(self.g), int(self.b)))
        self.set_color()


    def set_color(self):
        self.r_slider = ttk.Scale(self.frame_color, value=255, from_=0, to=255, command=self.bgUpdate, orient=tk.VERTICAL)
        self.r_slider.grid(row=1, column=0, rowspan=5)
        self.g_slider = ttk.Scale(self.frame_color, value=0, from_=0, to=255, command=self.bgUpdate, orient=tk.VERTICAL)
        self.g_slider.grid(row=1, column=1, rowspan=5)
        self.b_slider = ttk.Scale(self.frame_color, value=0, from_=0, to=255, command=self.bgUpdate, orient=tk.VERTICAL)
        self.b_slider.grid(row=1, column=2, rowspan=5)
        self.color_button = ttk.Button(self.frame_color, text="确认", command=self.save_color)
        self.color_button.grid(row=1, column=3, rowspan=3)

    def bgUpdate(self, x):
        self.r = int(self.r_slider.get())  # 获取数据
        self.g = int(self.g_slider.get())
        self.b = int(self.b_slider.get())
        myColor = "#%02x%02x%02x" % (self.r, self.g, self.b)  # 十六进制化
        self.root.configure(background=myColor)

    def save_color(self):
        self.app.toucolors[self.num] = [self.r,self.g,self.b]
        self.app.creat_tou()
        self.root.destroy()



class Application(tk.Frame):
    def __init__(self, root=None):
        super().__init__(root)
        # 界面配置
        self.root = root
        self.set_size()
        self.frame = ttk.Frame(root, padding="3 3 12 12")
        self.frame.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        # 存储变量
        self.zhaofiles = [tk.StringVar()]
        self.toufiles = [tk.StringVar()]
        self.toucolors = [[255,0,0]]
        self.splitnum = tk.StringVar()
        self.splitnum.set('10')
        self.splitword = tk.StringVar()
        self.splitword.set('。|:|：|,|，')
        self.log = tk.StringVar()
        # 创建控件
        self.create_widget()

    def set_size(self):
        self.root.title('投标文件比对工具')

    def create_widget(self):
        # 标题
        # self.label_title = ttk.Label(self.frame, text='投标文件比对工具', width=80)
        # self.label_title.grid(row=0, sticky=(tk.W, tk.E), columnspan=16)
        ttk.Button(self.frame, text="说明", command=self.readme).grid(column=0, row=1, columnspan=1)
        # 招标文件
        ttk.Label(self.frame, text="招标文件选择").grid(column=0, row=1, columnspan=19)
        self.zhaoAll = []
        self.creat_zhao()
        # 投标文件
        ttk.Label(self.frame, text="投标文件选择").grid(column=0, row=5, columnspan=19)
        self.touAll = []
        self.creat_tou()
        # 设置
        ttk.Label(self.frame, text="最小检测长度:").grid(column=0, row=20, columnspan=4)
        ttk.Entry(self.frame, textvariable=self.splitnum).grid(column=4, row=20,columnspan=4)
        ttk.Label(self.frame, text="分割词（正则）:").grid(column=8, row=20, columnspan=4)
        ttk.Entry(self.frame, textvariable=self.splitword).grid(column=12, row=20,columnspan=4)

        self.startbutton = ttk.Button(self.frame, text="对比", command=self.work)
        self.startbutton.grid(column=16, row=20,columnspan=3)

        # 日志
        self.text = tk.Text(self.frame, height=10)
        self.text.grid(column=0,row=21,columnspan=19)


    def print(self, x):
        self.text.insert('end', x+'\n')

    def stopwork(self):
        pass

    def work(self):
        try:
            if self.mt_working.is_alive():
                self.print("【程序已在进行中，请等待完成。】")
                return
        except:
            pass
        self.text.delete("1.0", 'end')
        zhaofiles, toufiles = [], []
        for file in self.zhaofiles:
            if not os.path.isfile(file.get()):
                self.print("招标文件不存在")
                return
            zhaofiles.append(file.get())
        for file in self.toufiles:
            if not os.path.isfile(file.get()):
                self.print("招标文件不存在")
                return
            toufiles.append(file.get())
        try:
            limitnum = int(self.splitnum.get())
        except:
            self.print('请输入数字')
        splitword = self.splitword.get()
        colors = self.toucolors
        # compare(zhaofiles, toufiles, limitnum, splitword, colors, self.print)
        self.mt_working = mt.Thread(target=compare, args=(zhaofiles, toufiles, limitnum, splitword, colors, self.print))
        self.mt_working.start()
        # root.after(1000, lambda: self.step_commit())
        self.mt_working.is_alive()

    def creat_zhao(self):
        for i in range(len(self.zhaoAll)):
            self.zhaoAll[i].destroy()
        for i in range(len(self.zhaofiles)):
            self.zhaoAll.append(ttk.Entry(self.frame, textvariable=self.zhaofiles[i]))
            self.zhaoAll[-1].grid(row=i+2, column=0, columnspan=16, sticky=tk.E + tk.W)
            self.zhaoAll.append(ttk.Button(self.frame, text="...", command=lambda i=i:self.selectFile('zhao', i)))
            self.zhaoAll[-1].grid(row=i+2, column=16)
            self.zhaoAll.append(ttk.Button(self.frame, text="—", command=lambda i=i:self.removeFile('zhao', i)))
            self.zhaoAll[-1].grid(row=i+2, column=18)
        if len(self.zhaofiles)<5:
            self.zhaoAll.append(ttk.Button(self.frame, text="+", command=lambda: self.addFile('zhao')))
            self.zhaoAll[-1].grid(row=i + 3, column=18)

    def creat_tou(self):
        for i in range(len(self.touAll)):
            self.touAll[i].destroy()
        for i in range(len(self.toufiles)):
            self.touAll.append(ttk.Entry(self.frame, textvariable=self.toufiles[i]))
            self.touAll[-1].grid(row=i+6, column=0, columnspan=16, sticky=tk.E + tk.W)
            self.touAll.append(ttk.Button(self.frame, text="...", command=lambda i=i:self.selectFile('tou', i)))
            self.touAll[-1].grid(row=i+6, column=16)
            self.touAll.append(tk.Button(self.frame, text="C", command=lambda i=i: self.set_color(i), bg="#%02x%02x%02x" % (int(self.toucolors[i][0]), int(self.toucolors[i][1]), int(self.toucolors[i][2]))))
            self.touAll[-1].grid(row=i+6, column=17)
            self.touAll.append(ttk.Button(self.frame, text="—", command=lambda i=i:self.removeFile('tou', i)))
            self.touAll[-1].grid(row=i+6, column=18)
        if len(self.toufiles)<10:
            self.touAll.append(ttk.Button(self.frame, text="+", command=lambda: self.addFile('tou')))
            self.touAll[-1].grid(row=i + 7, column=18)


    def selectFile(self, type, num):
        path = tkf.askopenfilename(filetypes =[("DOCX", ".docx")])
        print(path)
        if type=='zhao':
            self.zhaofiles[num].set(path)
        if type=='tou':
            self.toufiles[num].set(path)

    def addFile(self, type):
        if type=='zhao':
            self.zhaofiles.append(tk.StringVar())
            self.creat_zhao()
        if type=='tou':
            self.toufiles.append(tk.StringVar())
            self.toucolors.append([255,0,0])
            self.creat_tou()

    def removeFile(self, type, num):
        if type=='zhao':
            zf = []
            for j in range(len(self.zhaofiles)):
                if j!=num:
                    zf.append(self.zhaofiles[j])
            if len(zf)==0:
                zf = [tk.StringVar()]
            self.zhaofiles = zf
            self.creat_zhao()
        if type=='tou':
            zf = []
            for j in range(len(self.toufiles)):
                if j!=num:
                    zf.append(self.toufiles[j])
            if len(zf)==0:
                zf = [tk.StringVar()]
            self.toufiles = zf
            self.creat_tou()

    def set_color(self, num):
        r = tk.Tk()
        colorset = ColorSet(self, num, r)
        sv_ttk.use_light_theme()
        r.mainloop()

    def readme(self):
        r = tk.Tk()
        Readme(r)
        r.mainloop()



if __name__ == '__main__':
    root = tk.Tk()
    app = Application(root)
    # app = ColorSet(1,2,root)
    # button = ttk.Button(root, text="Click me!")
    # button.pack()
    sv_ttk.use_light_theme()
    root.mainloop()
