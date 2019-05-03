# -*- coding: utf-8 -*-
from tkinter import *
from tkinter import ttk
import tkinter, tkinter.constants, tkinter.filedialog
from tkinter.font import Font
from rand_class import *
from load_data import *
from write_data import *
import re

orig_data = []
result_data = []
th_tab = 0

class NestedPanesDemo(ttk.Frame):

    def convert65536(self, s):
        #Converts a string with out-of-range characters in it into a string with codes in it.
        l=list(s);
        i=0;
        while i<len(l):
            o=ord(l[i]);
            if o>65535:
                l[i]="{"+str(o)+"ū}";
            i+=1;
        return "".join(l);

    def parse65536(self, match):
        #This is a regular expression method used for substitutions in convert65536back()
        text=int(match.group()[1:-2]);
        if text>65535:
            return chr(text);
        else:
            return "ᗍ"+str(text)+"ūᗍ";
    def convert65536back(self, s):
        #Converts a string with codes in it into a string with out-of-range characters in it
        while re.search(r"{\d\d\d\d\d+ū}", s)!=None:
            s=re.sub(r"{\d\d\d\d\d+ū}", self.parse65536, s);
        s=re.sub(r"ᗍ(\d\d\d\d\d+)ūᗍ", r"{\1ū}", s);
        return s;

    def __init__(self, isapp=True, name='class_grouping'):
        ttk.Frame.__init__(self, name=name)
        self.pack(expand=Y, fill=BOTH)
        self.master.title('雲林縣常態編班程式')
        self.isapp = isapp
        self._create_widgets()

    def _create_widgets(self):

        self._create_group_panel()

    def _create_group_panel(self):
        groupPanel = ttk.Frame(self, name='group')
        groupPanel.pack(side=TOP, fill=BOTH, expand=Y)

        self._create_wnd_struct(groupPanel)
        self._button_pane()
        self._status_pane()
        self._grouping_status_pane()

    def _create_wnd_struct(self, parent):
        outer = ttk.PanedWindow(parent, orient=VERTICAL, name='outer')

        top = ttk.PanedWindow(outer, orient=HORIZONTAL, name='top')
        bot = ttk.PanedWindow(outer, orient=HORIZONTAL, name='bot')
        tl = ttk.LabelFrame(top, text='學校資訊', padding=3, name='tleft',width=300, height=140)
        tlm = ttk.LabelFrame(top, text='編班狀態', padding=3, name='tmid',width=300, height=140)
        tr = ttk.LabelFrame(top, text='操作', padding=3, name='tright', width=300, height=140)
        top.add(tl)
        top.add(tlm)
        top.add(tr)
        outer.pack(side=LEFT, expand=Y, fill=BOTH)
        outer.add(top)

        outer.pack(side=RIGHT, expand=Y, fill=BOTH)
        outer.add(bot)

        br = ttk.LabelFrame(bot, text='導師名單', padding=3, name='bright', width=240, height=420)
        bot.add(br)

    def _button_pane(self):
        # create and add button
        tright = self.nametowidget('group.outer.top.tright')
        bo = ttk.Button(tright, text='開檔', command=self.askopenfilename)
        bo.pack(expand=0, padx=2, pady=3)
        bo = ttk.Button(tright, text='編班', command=self.group_func)
        bo.pack(expand=0, padx=2, pady=3)
        bo = ttk.Button(tright, text='存檔', command=self.write_func)
        bo.pack(expand=0, padx=2, pady=3)

    def _grouping_status_pane(self):
        tlm = self.nametowidget('group.outer.top.tmid')
        midlabel = ttk.Frame(tlm)
        midlabel.pack()
        self.grouping_status = StringVar()
        self.grouping_status.set("請選擇學校")
        self.label_grouping_status = Label(midlabel, textvariable = self.grouping_status, fg = "blue", font="16")
        self.label_grouping_status.pack()

        botr = self.nametowidget('group.outer.bot.bright')
        teacherframe = ttk.Frame(botr)
        teacherframe.pack(side=TOP, fill=BOTH, expand=Y)
        self._create_teacher_treeview(teacherframe)

    def _status_pane(self):
        tleft = self.nametowidget('group.outer.top.tleft')
        leftlabel = ttk.Frame(tleft)
        leftlabel.pack()
        self.label_teacher_string = StringVar()
        self.label_teacher_string.set("導師人數:")
        self.label_teacher_number= Label(leftlabel, textvariable = self.label_teacher_string)
        self.label_teacher_number.pack()

    def _status_school_update(self):
        global orig_data
        self.teacher_number = len(orig_data)
        self.label_teacher_string.set("導師人數:"+str(self.teacher_number))

    def askopenfilename(self):
        global orig_data
        global result_data
        self.filename = tkinter.filedialog.askopenfilename()
        print ("open filename : %s" %self.filename)
# clean all data
        del orig_data[:]
        del result_data[:]
        for i in self.teacher_tree.get_children():
            self.teacher_tree.delete(i)
# clean all data
        orig_data = load_data(self.filename)
        tmp = self.filename.replace(".xlsx", "")
        schoolname = re.search(r'.*\d+(.*)$', tmp).group(1)
        self.grouping_status.set("讀取"+ schoolname +"教師資料")
        self._status_school_update()
        self._load_teacher_data()


    def group_func(self):
        print ("randoming!!")
        global orig_data
        global result_data
        result_data = rand_class(orig_data)
        self.grouping_status.set("教師隨機編定班級結束 請存檔")
        for i in self.teacher_tree.get_children():
            self.teacher_tree.delete(i)
        self._load_teacher_data()


    def write_func(self):
        global result_data
        global orig_data
        writefile(result_data, self.filename)
        print ("writing file!!")
        self.grouping_status.set("存檔結束")
# clean all data
        del orig_data[:]
        del result_data[:]
        for i in self.teacher_tree.get_children():
            self.teacher_tree.delete(i)
        self.label_teacher_string.set("導師人數:")

    def _create_teacher_treeview(self, parent):
        f = ttk.Frame(parent)
        f.pack(side=TOP, fill=BOTH, expand=Y)

        # create the tree and scrollbars
        self.teacherCols = ('序號', '年級', '性別', '姓名', '編訂班別')
        self.teacher_tree = ttk.Treeview(columns=self.teacherCols,
                                 show = 'headings')

        ysb = ttk.Scrollbar(orient=VERTICAL, command= self.teacher_tree.yview)
        xsb = ttk.Scrollbar(orient=HORIZONTAL, command= self.teacher_tree.xview)
        self.teacher_tree['yscroll'] = ysb.set
        self.teacher_tree['xscroll'] = xsb.set

        # add teacher_tree and scrollbars to frame
        self.teacher_tree.grid(in_=f, row=0, column=0, sticky=NSEW)
        ysb.grid(in_=f, row=0, column=1, sticky=NS)
        xsb.grid(in_=f, row=1, column=0, sticky=EW)

        # set frame resize priorities
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

    def _load_teacher_data(self):
      # configure column headings
        for c in self.teacherCols:
            self.teacher_tree.heading(c, text=c.title(),
                              command=lambda c=c: self._column_sort(c, MCListDemo.SortDir))
            self.teacher_tree.column(c, minwidth=100, width=Font().measure(c.title()))

        # add data to the teacher_tree
        for item in orig_data:
            item[1] = self.convert65536(''.join(item[1]))
            self.teacher_tree.insert('', 'end', values=(item[0],item[1],item[2], item[3], item[4]))
            item[1] = self.convert65536back(''.join(item[1]))


if __name__ == '__main__':
    root = Tk()
    root.call('encoding','system','unicode')
    root.minsize(width=1000, height=600)
    root.resizable(width=False, height=False)
    NestedPanesDemo().mainloop()
