#!/usr/bin/env python3
"""
Ver 0.0.1 - First version
Ver 0.0.2 - add split line on txt log out for user
            Remove comment col first
            Change to check both of all txt and csv fail log
Ver 0.0.3 - Change get file name and show rule for match factory store log new rule
Ver 0.0.4 - Fix csv no file will make form error issue
Ver 0.0.5 - implement user input specif sn to get relate csv and txt log
Ver 0.0.6 - Add more fail condition
            Remove \n on user input sn
            Save html to excel format
Ver 0.0.7 - Fix Calculation FA fail range issue
            Fix Calculation Count fail range issue
Ver 0.0.8 - Add red color for Count and FA fail number in fail log report
"""

from tkinter import *
from tkinter.ttk import *
from tkinter.font import Font
import tkinter.messagebox
import re
import configparser
import os
import shutil
from collections import OrderedDict
from datetime import datetime, timedelta
# import pandas

# from tkinter.scrolledtext import ScrolledText
# from time import sleep
# from tkinter.commondialog import Dialog
from enum import Enum

version = "v0.0.8"
TITLE = "LogParser - " + version
SETTING_NAME = "Settings.ini"
html_template_need_repeat = \
'    <tr>\n'\
'        <th align="center">{{sn}}</th>\n'\
'        <th align="center"><a href="{{csv_url}}">{{csv_filename}}</a></th>\n'\
'        <td align="center">{{csv_fail_log}}</td>\n'\
'        <th align="center"><a href="{{txt_url}}">{{txt_file_name}}</a></th>\n\n'\
'        <td align="center">{{txt_fail_log}}</td>\n'\
'        <!--<td align="center">{{comment}}</td>-->\n'\
'    </tr>\n'\
'{{next_item}}'


sub_database_name = "SubList.sdb"

backup_folder_name = "backfile"
subpath = ""  # SUB file path, read from Settings.ini
subfiletype_list = ""  # SUB file type, read from Settings.ini, ex: *.ssa, *.ass

progress_idle_txt = ""

help_text = \
    "有問題找請william\n"\
    "william_liu@compal.com\n"\
    "ext: #18608\n"\

class error_Type(Enum):
    NORMAL = 'NORMAL'  # define normal state
    FILE_ERROR = 'FILE_RW_ERROR'  # define file o/r/w error type


class replace_Sub_Gui(Frame):
    def __init__(self, master=None, subfilepath_ini=None, subfiletype_ini=None,
                 csvfilepath_ini=None, txtfilepath_ini=None, help_text=None):

        Frame.__init__(self, master)
        self.master = master
        self.subfiletype_list_ini = subfiletype_ini
        self.subpath_ini = subfilepath_ini
        self.csvpath_ini = csvfilepath_ini
        self.txtpath_ini = txtfilepath_ini
        self.user_input_csvpath = ""
        self.user_input_txtpath = ""
        self.help_text = help_text
        self.user_input_path = ""
        self.user_input_type = ""

        root.bind('<Key-Return>', self.press_key_enter)

        self.create_widgets()

    def create_widgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.style.configure('Tlog_frame.TLabelframe', font=('iLiHei', 10))
        self.style.configure('Tlog_frame.TLabelframe.Label', font=('iLiHei', 10))
        self.log_frame = LabelFrame(self.top, text='LOG', style='Tlog_frame.TLabelframe')
        self.log_frame.place(relx=0.01, rely=0.283, relwidth=0.973, relheight=0.708)

        self.style.configure('Tuser_input_frame.TLabelframe', font=('iLiHei', 10))
        self.style.configure('Tuser_input_frame.TLabelframe.Label', font=('iLiHei', 10))
        self.user_input_frame = LabelFrame(self.top, text='輸入', style='Tuser_input_frame.TLabelframe')
        self.user_input_frame.place(relx=0.01, rely=0.011, relwidth=0.973, relheight=0.262)

        self.VScroll1 = Scrollbar(self.log_frame, orient='vertical')
        self.VScroll1.place(relx=0.967, rely=0.010, relwidth=0.022, relheight=0.936)

        self.HScroll1 = Scrollbar(self.log_frame, orient='horizontal')
        self.HScroll1.place(relx=0.01, rely=0.940, relwidth=0.958, relheight=0.055)

        self.log_txtFont = Font(font=('iLiHei', 10))
        self.log_txt = Text(self.log_frame, wrap='none', xscrollcommand=self.HScroll1.set, yscrollcommand=self.VScroll1.set, font=self.log_txtFont)
        self.log_txt.place(relx=0.01, rely=0.010, relwidth=0.958, relheight=0.936)
        # self.log_txt.insert('1.0', '')
        self.HScroll1['command'] = self.log_txt.xview
        self.VScroll1['command'] = self.log_txt.yview

        self.style.configure('Thelp_button.TButton', font=('iLiHei', 9))
        self.help_button = Button(self.user_input_frame, text='Help', command=self.print_about, style='Thelp_button.TButton')
        self.help_button.place(relx=0.460, rely=0.788, relwidth=0.105, relheight=0.200)

        self.style.configure('Tstart_button.TButton', font=('iLiHei', 9))
        self.start_button = Button(self.user_input_frame, text='Start', command=self.analysis_log_and_gen_report, style='Tstart_button.TButton')
        self.start_button.place(relx=0.250, rely=0.788, relwidth=0.105, relheight=0.200)

        self.csv_path_entryVar = StringVar(value=self.csvpath_ini)
        self.csv_path_entry = Entry(self.user_input_frame, textvariable=self.csv_path_entryVar, font=('iLiHei', 10))
        self.csv_path_entry.place(relx=0.01, rely=0.180, relwidth=0.80, relheight=0.180)

        self.txt_path_entryVar = StringVar(value=self.txtpath_ini)
        self.txt_path_entry = Entry(self.user_input_frame, textvariable=self.txt_path_entryVar, font=('iLiHei', 10))
        self.txt_path_entry.place(relx=0.01, rely=0.520, relwidth=0.80, relheight=0.190)

        self.style.configure('Tversion_label.TLabel', anchor='e', font=('iLiHei', 9))
        self.version_label = Label(self.user_input_frame, text=version, state='disable', style='Tversion_label.TLabel')
        self.version_label.place(relx=0.843, rely=0.87, relwidth=0.147, relheight=0.13)

        self.style.configure('Tversion_state.TLabel', anchor='w', font=('iLiHei', 9))
        self.version_state = Label(self.user_input_frame, text=progress_idle_txt, style='Tversion_state.TLabel')
        self.version_state.place(relx=0.01, rely=0.87, relwidth=0.116, relheight=0.13)

        self.style.configure('Ttxt_path_label.TLabel', anchor='w', font=('iLiHei', 10))
        self.txt_path_label = Label(self.user_input_frame, text='SN序號, 請用;隔開, 例如:731976203;731972421;731974348', style='Ttxt_path_label.TLabel')
        self.txt_path_label.place(relx=0.01, rely=0.380, relwidth=0.600, relheight=0.13)

        self.style.configure('Tcsv_path_label.TLabel', anchor='w', font=('iLiHei', 10))
        self.csv_path_label = Label(self.user_input_frame, text='log檔根目錄路徑', style='Tcsv_path_label.TLabel')
        self.csv_path_label.place(relx=0.01, rely=0.010, relwidth=0.200, relheight=0.166)

        self.log_txt.tag_config("error", foreground="#CC0000")
        self.log_txt.tag_config("info", foreground="#008800")
        self.log_txt.tag_config("info2", foreground="#404040")

        self.update_idletasks()

    def press_key_enter(self, event=None):
        self.analysis_log_and_gen_report()

    # def convert_clipboard(self):
    #     # -----Clear text widge for log-----
    #     # self.log_txt.config(state="normal")
    #     # self.log_txt.delete('1.0', END)
    #     # self.log_txt.config(state="disable")
    #
    #     clip_content_lv = self.clipboard_get()
    #     self.clipboard_clear()
    #     conv_ls = self.get_user_conv_type()
    #     # clip_content_lv = langconver.convert_lang_select(clip_content_lv, 's2t')
    #     # print(conv_ls)
    #     clip_content_lv = langconver.convert_lang_select(clip_content_lv, conv_ls)
    #     self.clipboard_append(clip_content_lv)
    #
    #     self.setlog("剪貼簿轉換完成!", 'info')

    def print_about(self):
        tkinter.messagebox.showinfo("About", self.help_text)

    def setlog(self, string, level=None):
        self.log_txt.config(state="normal")

        if (level != 'error') and (level != 'info') and (level != 'info2'):
            level = ""

        self.log_txt.insert(INSERT, "%s\n" % string, level)
        # -----scroll to end of text widge-----
        self.log_txt.see(END)
        self.update_idletasks()

        self.log_txt.config(state="disabled")

    def setlog_large(self, string, level=None):
        self.log_txt.insert(INSERT, "%s\n" % string, level)
        # -----scroll to end of text widge-----
        self.log_txt.see(END)
        self.update_idletasks()

    def read_config(self, filename, section, key):
        try:
            config_lh = configparser.ConfigParser()
            file_ini_lh = open(filename, 'r', encoding='utf16')
            config_lh.read_file(file_ini_lh)
            file_ini_lh.close()
            return config_lh.get(section, key)
        except:
            self.setlog("Error! 讀取ini設定檔發生錯誤! "
                        "請在" + TITLE + "目錄下使用UTF-16格式建立 " + filename, 'error')
            return error_Type.FILE_ERROR.value

    def write_config(self, filename, sections, key, value):
        try:
            config_lh = configparser.ConfigParser()
            file_ini_lh = open(filename, 'r', encoding='utf16')
            config_lh.read_file(file_ini_lh)
            file_ini_lh.close()

            file_ini_lh = open(filename, 'w', encoding='utf16')
            config_lh.set(sections, key, value)
            config_lh.write(file_ini_lh)
            file_ini_lh.close()
        except Exception as ex:
            self.setlog("Error! 寫入ini設定檔發生錯誤! "
                        "請在" + TITLE + "目錄下使用UTF-16格式建立 " +filename, 'error')
            return error_Type.FILE_ERROR.value

    def get_file_list(self, file_path, file_type):
        file_list_ll = []
        # tkinter.messagebox.showerror("info", "into get file list")

        for root, dirs, files in os.walk(file_path):
            for file in files:
                if file.endswith(file_type):
                    print(os.path.join(root, file))
                    file_list_ll.append(os.path.join(root, file))

        return file_list_ll

    def get_csv_file_list_and_sn_to_dic(self, file_path, file_type):
        sn_and_filepath_odic = OrderedDict()
        # tkinter.messagebox.showerror("info", "into get file list")

        for root, dirs, files in os.walk(file_path):
            for file in files:
                if file.endswith(file_type):
                    re_h = re.match(r'(.+)_(.+)', file)
                    # subdata_dic_ld[re_h.group(1)] = re_h.group(2)
                    sn_and_filepath_odic[re_h.group(1)] = os.path.join(root, file)
                    # mapping_orisub_and_video_odic[i] = videofile_list_ll[count]
                    # print(os.path.join(root, file))
                    # file_list_ll.append(os.path.join(root, file))

        return sn_and_filepath_odic

    def find_file_error_and_store_to_dic(self, key, filepath):
        sn_and_error_odic = OrderedDict()
        error_ll = []
        find_error = False
        with open(filepath, "r") as ins:
            error_ll.clear()
            for line in ins:
                if line.find('FAIL') > -1 or line.find('fail') > -1 or line.find('Fail') > -1 \
                        or line.find('no devices found') > -1 or line.find('device offline') > -1:
                    find_error = True
                    error_ll.append(line)
        sn_and_error_odic[key] = error_ll
                    # print(line)

        # for key, error in sn_and_error_odic.items():
        #     print('-----%s-----' % key)
        #     for i in error:
        #         print(i)

        return find_error, sn_and_error_odic

    def find_txtfile_error_and_store_to_dic(self, key, filepath):
        sn_and_error_odic = OrderedDict()
        error_ll = []
        find_error = False
        with open(filepath, "r") as ins:
            error_ll.clear()
            for line in ins:
                count_fail = False
                fa_fail = False
                fa_value_ll = []
                # a = re.match('Motion Count:', line)
                if line.find('Motion Count: ') > -1:
                    # print(line)
                    re_h = re.match('.*SPEC\((.*)\)', line)
                    range_value = re_h.group(1).split('~')
                    count_minimum = range_value[0]
                    count_maximum = range_value[1]
                    re_h = re.match('.*Motion Count: (\d+),', line)
                    if re_h:
                        count_value = re_h.group(1)
                        count_value = re.sub(' ', '', count_value)
                        # print(count_maximum, count_minimum, count_value)
                        if (count_value > count_maximum) or (count_value < count_minimum):
                            count_fail = True
                        # print('count out of range')
                if line.find('Motion FA: ') > -1:
                    re_h = re.match('.*SPEC\((.*)\)', line)
                    range_value = re_h.group(1).split('~')
                    fa_minimum = range_value[0]
                    fa_maximum = range_value[1]
                    re_h = re.match('.*Motion FA: (.*)SPEC\(', line)
                    if re_h:
                        temp_str = re.sub(' ', '', re_h.group(1))
                        temp_str = re.sub(',,', '', temp_str)
                        fa_value_ll = temp_str.split(',')
                        for fa_value in fa_value_ll:
                            if (fa_value > fa_maximum) or (fa_value < fa_minimum):
                                fa_fail = True

                if line.find('FAIL') > -1 or line.find('fail') > -1 or line.find('Fail') > -1 \
                        or line.find('no devices found') > -1 or line.find('device offline') > -1\
                        or count_fail or fa_fail:
                    # print('william:%s' % line)
                    find_error = True
                    error_ll.append(line)

        sn_and_error_odic[key] = error_ll


        # for key, error in sn_and_error_odic.items():
        #     print('-----%s-----' % key)
        #     for i in error:
        #         print(i)

        return find_error, sn_and_error_odic


    # def get_txt_file_list_match_keyword_to_dic(self, file_path, file_type, keyword):
    def get_file_list_store_to_dic(self, all_txt_path):
        sn_and_filepath_odic = OrderedDict()
        # all_txt_path = []
        txt_list_ll = []
        find_txt_file = False

        for file in all_txt_path:
            file_name = os.path.split(file)
            file_name_and_ext = os.path.splitext(file_name[1])
            if file_name_and_ext[1] == r'.csv':
                re_h = re.match(r'(.+)_calib', file_name_and_ext[0])
                sn_and_filepath_odic[re_h.group(1)] = file
            else:
                sn_and_filepath_odic[file_name_and_ext[0]] = file

        # for k, v in sn_and_filepath_odic.items():
        #     print(k)
        #     print(v)

        return sn_and_filepath_odic


    def store_file_to_folder(self, file, back_folder):
        shutil.copy2(file, back_folder)

    def change_html_color_count_fa_fail_number(self, one_line_fail_text):
        fa_value_ll = []
        # count_fail = False
        # fa_fail = False
        input_str = str(one_line_fail_text)
        # a = re.match('Motion Count:', line)
        if input_str.find('Motion Count: ') > -1:
            # print(input_str)
            re_h = re.match('.*SPEC\((.*)\)', input_str)
            range_value = re_h.group(1).split('~')
            count_minimum = range_value[0]
            count_maximum = range_value[1]
            re_h = re.match('.*Motion Count: (\d+),', input_str)
            if re_h:
                count_value = re_h.group(1)
                count_value = re.sub(' ', '', count_value)
                # print(count_value)
                # print(count_maximum, count_minimum, count_value)
                if (count_value > count_maximum) or (count_value < count_minimum):
                    input_str = input_str.replace('%s%s' % (count_value, ','), '%s%s%s%s' % (r'<font color="red">', count_value, ',', r'</font>'))

            return input_str

        elif input_str.find('Motion FA: ') > -1:
            # print(input_str)
            re_h = re.match('.*SPEC\((.*)\)', input_str)
            range_value = re_h.group(1).split('~')
            fa_minimum = range_value[0]
            fa_maximum = range_value[1]
            re_h = re.match('.*Motion FA: (.*) SPEC\(', input_str)
            if re_h:
                temp_str = re.sub(' ', '', re_h.group(1))
                temp_str = re.sub(',,', '', temp_str)
                fa_value_ll = temp_str.split(',')
                for fa_value in fa_value_ll:
                    if (fa_value > fa_maximum) or (fa_value < fa_minimum):
                        input_str = input_str.replace('%s%s' % (fa_value, ','), '%s%s%s%s' % (r'<font color="red">', fa_value, ',', r'</font>'))
            return input_str

        else:
            return one_line_fail_text

    def analysis_log_and_gen_report(self):
        sn_and_csverror_odic = OrderedDict()
        local_csv_path_odic = OrderedDict()
        txt_sn_and_filepath_odic = OrderedDict()
        sn_and_txterror_odic = OrderedDict()
        w_file_stat_lv = error_Type.NORMAL.value

        # -----Clear text widge for log-----
        self.log_txt.config(state="normal")
        self.log_txt.delete('1.0', END)
        self.log_txt.config(state="disable")

        # -----Get user input csv path-----
        self.user_input_csvpath = self.csv_path_entry.get()
        # -----Get user input txt path-----
        self.user_input_txtpath = self.txt_path_entry.get()
        # -----Check user input in GUI-----
        if self.user_input_csvpath == "" or self.user_input_txtpath == "":
            tkinter.messagebox.showinfo("message", "請輸入csv與txt路徑及sn")
            return
        if not os.path.exists(self.user_input_csvpath):
            tkinter.messagebox.showerror("Error", "路徑錯誤")
            return

        # -----get config ini file setting-----
        self.csvpath_ini = self.read_config(SETTING_NAME, 'Global', 'csvpath')
        self.txtpath_ini = self.read_config(SETTING_NAME, 'Global', 'txtpath')
        if self.csvpath_ini == error_Type.FILE_ERROR.value or self.txtpath_ini == error_Type.FILE_ERROR.value:
            tkinter.messagebox.showerror("Error",
                                         "錯誤! 讀取ini設定檔發生錯誤! "
                                         "請在" + TITLE + "目錄下使用UTF-16格式建立 " + SETTING_NAME)
            return

        # -----remove '\' or '/' in end of path string-----
        self.user_input_csvpath = re.sub(r"/$", '', self.user_input_csvpath)
        self.user_input_csvpath = re.sub(r"\\$", "", self.user_input_csvpath)
        self.user_input_txtpath = re.sub(r" ", '', self.user_input_txtpath)
        self.user_input_txtpath = re.sub(r"\n", "", self.user_input_txtpath)

        # -----Store user input path and type into Setting.ini config file-----
        if not self.user_input_csvpath == self.csvpath_ini:
            self.setlog("新的csv路徑設定寫入設定檔: " + SETTING_NAME, "info")
            # print("path not match, write new path to ini")
            w_file_stat_lv = self.write_config(SETTING_NAME,  'Global', 'csvpath', self.user_input_csvpath)
        if not self.user_input_txtpath == self.txtpath_ini:
            self.setlog("新的sn寫入設定檔: " + SETTING_NAME, "info")
            # print("type not match, write new type list to ini")
            w_file_stat_lv = self.write_config(SETTING_NAME, 'Global', 'txtpath', self.user_input_txtpath)
        if w_file_stat_lv == error_Type.FILE_ERROR.value:
            tkinter.messagebox.showerror("Error",
                                         "錯誤! 寫入ini設定檔發生錯誤! "
                                         "請在" + TITLE + "目錄下使用UTF-16格式建立 " + SETTING_NAME)
            return

        # store file folder name
        log_foldername = re.sub(r":", '.', str(datetime.utcnow() + timedelta(hours=8)))
        csv_foldername = ('%s\\%s\\%s\\' % (os.getcwd(), log_foldername, 'csv'))
        txt_foldername = ('%s\\%s\\%s\\' % (os.getcwd(), log_foldername, 'txt'))

        if not os.path.exists(csv_foldername):
            os.makedirs(csv_foldername)
        if not os.path.exists(txt_foldername):
            os.makedirs(txt_foldername)

        #     **************

        self.setlog("讀取所有txt檔", 'info')
        # # get all txt file that fail in csv log
        # for csv_key in sn_and_csverror_odic.keys():
        #     temp_odic = self.get_txt_file_list_match_keyword_to_dic(self.user_input_txtpath, '.txt', csv_key)
        #     txt_sn_and_filepath_odic.update(temp_odic)

        # get all txt file that have sn in csv file

        all_txt_path = []
        for root, dirs, files in os.walk(self.user_input_csvpath):
            for file in files:
                if file.endswith('.txt'):
                    all_txt_path.append(os.path.join(root, file))


        # full_path = r'C:\20170705_log\20170704_32A_E\FAIL\034427-731974736.txt'
        sn_ll = []
        sn_user_ll = self.user_input_txtpath.split(';')

        for input_sn in sn_user_ll:
            for full_path in tuple(all_txt_path):
                full_file_path = os.path.split(full_path)
                file_name_and_ext = os.path.splitext(full_file_path[1])
                file_name = file_name_and_ext[0]
                file_sn = re.sub('.*-', '', file_name)
                if input_sn == file_sn:
                    sn_ll.append(file_name)


        txt_sn_and_filepath_odic = self.get_file_list_store_to_dic(all_txt_path)

        if not txt_sn_and_filepath_odic:
            # csv file list is empty
            self.setlog("錯誤! 在指定的目錄中找不到txt檔案! 請確認輸入路徑", 'error')
            tkinter.messagebox.showwarning("Error", "錯誤! 在指定的目錄中找不到txt檔案! 請確認輸入路徑")
            return

        self.setlog("讀取所有csv檔", 'info')
        all_csv_path = []
        for root, dirs, files in os.walk(self.user_input_csvpath):
            for file in files:
                if file.endswith('.csv'):
                    all_csv_path.append(os.path.join(root, file))

        csv_sn_and_filepath_odic = self.get_file_list_store_to_dic(all_csv_path)
        if not csv_sn_and_filepath_odic:
            # csv file list is empty
            self.setlog("錯誤! 在指定的目錄中找不到csv檔案! 請確認輸入路徑", 'error')
            tkinter.messagebox.showwarning("Error", "錯誤! 在指定的目錄中找不到csv檔案! 請確認輸入路徑")
            return
        #
        # for k,v in csv_sn_and_filepath_odic.items():
        #     print(k)
        #     print(v)
        # return

        self.setlog("開始分析txt檔", 'info')
        # read txt content and find 'fail' keyword
        for txt_key, txt_path in txt_sn_and_filepath_odic.items():
            is_find_error, temp_odic = self.find_txtfile_error_and_store_to_dic(txt_key, txt_path)
            if is_find_error:
                sn_and_txterror_odic.update(temp_odic)

        # copy txt file that have fail log into log folder
        for i in sn_and_txterror_odic.keys():
            self.store_file_to_folder(txt_sn_and_filepath_odic[i], txt_foldername)

        # for k, v in sn_and_txterror_odic.items():
        #     print('=====%s=====' % k)
        #     for i in v:
        #         print(i)
        # return

        self.setlog("開始分析csv檔", 'info')
        # read all csv file and find 'fail' keyword
        for sn, full_path in csv_sn_and_filepath_odic.items():
            is_find_error, temp_odic = self.find_file_error_and_store_to_dic(sn, full_path)
            if is_find_error:
                sn_and_csverror_odic.update(temp_odic)

        # copy csv file that have fail log into log folder
        for i in sn_and_csverror_odic.keys():
            self.store_file_to_folder(csv_sn_and_filepath_odic[i], csv_foldername)

        # for k, v in sn_and_csverror_odic.items():
        #     print('=====%s=====' % k)
        #     for i in v:
        #         print(i)
        # return

        # ***************************************
        # copy temp.html to log folder
        self.store_file_to_folder('temp.html', log_foldername)
        # rename temp.html to result.html
        os.rename(('%s\\temp.html' % log_foldername), ('%s\\result.html' % log_foldername))
        html_full_path = ('%s\\result.html' % log_foldername)

        self.setlog("產生html 報表", 'info')
        # todo: for test, remove late, temp copy temp.html to test.html folder
        # self.store_file_to_folder('temp.html', 'test.html')
        # write to html
        try:
            # readfile_html_lh = open('test.html', 'r')
            readfile_html_lh = open(html_full_path, 'r')
            html_content = str(readfile_html_lh.read())
            readfile_html_lh.close()
        except Exception as ex:
            self.setlog("Error! 讀取html設定檔發生錯誤! "
                        "請確認" + TITLE + "目錄有temp.html且內容正確", 'error')
            return

        # html_content = html_content.replace('{{sn}}', '12345')
        # for sn, error in sn_and_csverror_odic.items(): # read all sn that fail in csv file
        # for sn in txt_sn_and_filepath_odic.keys():  # read all sn that has csv file
        for sn in sn_ll:  # read all sn that in user input sn list
            # if sn have error that in txt file
            sn_has_txt_error = False

            if (sn in sn_and_txterror_odic) or (sn in sn_and_csverror_odic):  # if sn have error that in txt file
                # print('-----%s-----' % k)
                # write sn
                if html_content.find('{{sn}}') == -1:
                    self.setlog("Error! 寫入html發生錯誤, 找不到{{sn}}標籤! ", 'error')
                    return
                html_content = html_content.replace('{{sn}}', sn)
                if sn in csv_sn_and_filepath_odic.keys():

                    # write csv file path to url
                    html_content = html_content.replace('{{csv_url}}', csv_sn_and_filepath_odic[sn])
                    # write csv file name
                    html_content = html_content.replace('{{csv_filename}}', ('%s%s' % (sn, '_calib.csv')))

                    if sn in sn_and_csverror_odic.keys():
                        # write csv error log to html
                        temp_error_log = ''
                        for fail_list in sn_and_csverror_odic[sn]:
                            if isinstance(fail_list, list):
                                for i in fail_list:
                                    temp_error_log += '%s%s' % (i, '<br>')
                            else:
                                temp_error_log += '%s%s' % (fail_list, '<br>')
                        html_content = html_content.replace('{{csv_fail_log}}', temp_error_log)
                    else:
                        html_content = html_content.replace('{{csv_fail_log}}', 'No error')
                        self.setlog("file: " + sn + '沒有csv fail log, 但有txt fail log', 'info')
                else:
                    # write csv file path to url
                    html_content = html_content.replace('{{csv_url}}', 'No file')
                    # write csv file name
                    html_content = html_content.replace('{{csv_filename}}', 'No file')
                    html_content = html_content.replace('{{csv_fail_log}}', 'No file')
                    self.setlog("file: " + sn + ' 有txt記錄檔但沒有csv檔', 'info')

                # write txt all file path to html
                temp_txt_filepath = ''
                if sn in sn_and_txterror_odic.keys():
                    # write txt file path to url
                    html_content = html_content.replace('{{txt_url}}', txt_sn_and_filepath_odic[sn])
                    # write txt file name
                    html_content = html_content.replace('{{txt_file_name}}', ('%s%s' % (sn, '.txt')))

                    temp_error_log = ''
                    for fail_list in sn_and_txterror_odic[sn]:
                        if isinstance(fail_list, list):
                            for i in fail_list:
                                test = i
                                test = self.change_html_color_count_fa_fail_number(test)
                                # temp_error_log += '%s%s' % (i, '<br>')
                                temp_error_log += '%s%s' % (test, '<br>')
                        else:
                            test = fail_list
                            test = self.change_html_color_count_fa_fail_number(test)
                            # temp_error_log += '%s%s' % (fail_list, '<br>')
                            temp_error_log += '%s%s' % (test, '<br>')
                        temp_error_log += '%s%s' % ('*******', '<br>')
                    html_content = html_content.replace('{{txt_fail_log}}', temp_error_log)
                else:
                    # write txt file path to url
                    html_content = html_content.replace('{{txt_url}}', txt_sn_and_filepath_odic[sn])
                    # write txt file name
                    html_content = html_content.replace('{{txt_file_name}}', ('%s%s' % (sn, '.txt')))
                    html_content = html_content.replace('{{txt_fail_log}}', 'No error')

                html_content = html_content.replace('{{next_item}}', html_template_need_repeat)

        html_content = html_content.replace('{{next_item}}', '')
        html_content = html_content.replace('{{sn}}', '')
        html_content = html_content.replace('{{csv_filename}}', '')
        html_content = html_content.replace('{{csv_fail_log}}', '')
        html_content = html_content.replace('{{txt_file_name}}', '')
        html_content = html_content.replace('{{txt_fail_log}}', '')
        html_content = html_content.replace('{{comment}}', '')


        try:
            # writhfile_html_lh = open('test.html', 'w')
            writhfile_html_lh = open(html_full_path, 'w')
            writhfile_html_lh.write(html_content)
            writhfile_html_lh.close()

        except Exception as ex:
            self.setlog("Error! 寫入html設定檔發生錯誤! "
                        "請確認" + TITLE + "目錄有temp.html且內容正確", 'error')
            return

        # self.setlog("已產生報表, 請打開 " + log_foldername + "\\result.html", 'info')
        self.setlog("已產生報表, 請打開 " + html_full_path, 'info')
        self.store_file_to_folder(html_full_path, '%s%s%s' % (log_foldername, '\\', 'result.xls'))
        # ================================================================================
        # with open(html_full_path, 'r') as f:
        #     df = pandas.read_html(f.read())
        # # print(df[0])
        # bb = pandas.ExcelWriter('%s%s%s' % (log_foldername, '\\', 'out.xlsx'))
        # # str(df[0]).replace('*******', '')
        # df[0].to_excel(bb)
        #
        # bb.close()

        # print(log_foldername)


        # for k, v in sn_and_csverror_odic.items():
        #     print('-----%s-----' % k)
        #     for i in v:
        #         print(i)

        # for i in csv_file_list:
        #     self.setlog(i, 'error')

        # if not sub_file_list:
        #     # convert file list is empty
        #     tkinter.messagebox.showwarning("Error", "錯誤! 在指定的目錄中找不到檔案! 請確認檔案路徑與類型")
        #     return

def check_all_file_status():
    if not os.path.exists(SETTING_NAME):
        return False
    if not os.path.exists('icons\\main.ico'):
        return False
    return True


if __name__ == '__main__':
    # -----MessageBox will create tkinter, so create correct setting tkinter first
    root = Tk()
    root.title(TITLE)
    root.iconbitmap('icons\\main.ico')

    SETTING_NAME = "%s\\%s" % (os.getcwd(), SETTING_NAME)
    # sub_database_name = "%s\\%s" % (os.getcwd(), sub_database_name)

    if not check_all_file_status():
        tkinter.messagebox.showerror("Error", "遺失必要檔案! \n\n請確認目錄有以下檔案存在, 或 "
                                              "重新安裝: " + TITLE + "\n"
                                              "1. " + SETTING_NAME + "\n"
                                              "2. icons\\main.ico")
        sys.exit(0)

    try:
        # -----Get setting from Settings.ini-----
        file_ini_h = open(SETTING_NAME, encoding='utf16')
        config_h = configparser.ConfigParser()
        config_h.read_file(file_ini_h)
        file_ini_h.close()
        subpath = config_h.get('Global', 'subpath')
        subfiletype_list = config_h.get('Global', 'subtype')
        csvpath = config_h.get('Global', 'csvpath')
        txtpath = config_h.get('Global', 'txtpath')
        config_h.clear()
    except:
        tkinter.messagebox.showerror("Error",
                                     "讀取設定檔 " + SETTING_NAME + " 錯誤!\n"
                                     "請確認檔案格式為UTF-16 (unicode format) 或重新安裝" + TITLE + "\n")
        sys.exit(0)

    # -----Start GUI class-----
    root.geometry('784x614')
    app = replace_Sub_Gui(master=root, subfilepath_ini=subpath, subfiletype_ini=subfiletype_list,
                          csvfilepath_ini=csvpath, txtfilepath_ini=txtpath, help_text=help_text)
    # -----Start main loop-----
    app.mainloop()
