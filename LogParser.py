# -*- coding: utf-8 -*-
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
Ver 0.0.9 - Add fixture col
            Add more detail in excel form
            Add more user input sn method
Ver 0.1.0 - Add comment col and classify fail log type
Ver 0.1.1 - keep user input sn order
Ver 0.1.2 - modify fail string
Ver 0.1.3 - Add more comment case to 15 and more USB disconect case
Ver 0.1.4 - Add more comment case to 16
            don't read/write temp_txt for speed up
Ver 0.1.4 - Modify comment string
Ver 0.1.5 - Modify comment string
Ver 0.1.6 - Add more comment case to 17
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

version = "v0.1.6"
TITLE = "LogParser - " + version
SETTING_NAME = "Settings.ini"
html_template_need_repeat = \
'    <tr>\n'\
'        <th align="center">{{sn}}</th>\n'\
'        <th align="center">{{fixture}}</th>\n'\
'        <th align="center"><a href="{{csv_url}}">{{csv_filename}}</a></th>\n'\
'        <td align="center">{{csv_fail_log}}</td>\n'\
'        <th align="center"><a href="{{txt_url}}">{{txt_file_name}}</a></th>\n\n'\
'        <td align="center">{{txt_fail_log}}</td>\n'\
'        <!--<td align="center">{{comment}}</td>-->\n'\
'    </tr>\n'\
'{{next_item}}'


form_template_need_repeat = \
'    <tr>\n'\
'        <th align="center">{{sn}}</th>\n'\
'        <th align="center">{{fixture}}</th>\n'\
'        <th align="center"><a href="{{csv_url}}">{{csv_filename}}</a></th>\n'\
'        <td align="center">{{csv_fail_log}}</td>\n'\
'        <th align="center"><a href="{{txt_url}}">{{txt_file_name}}</a></th>\n\n'\
'        <td align="center">{{txt_fail_log}}</td>\n'\
'        <td align="center">{{comment}}</td>\n'\
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
        self.txt_path_label = Label(self.user_input_frame, text='SN序號, 請用;或換行符隔開, 例如:731976203;731972421;731974348', style='Ttxt_path_label.TLabel')
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

    def check_button_clib_fail(self, filecontect):
        release_press_ll = []
        release_value_ll = []
        release_log_ll = []
        release_press_log_ll = []
        release_press_value_odic = OrderedDict()
        temp_release_press_log_odic = OrderedDict()
        release_press_log_odic = OrderedDict()
        is_button_clib_fail_type = -1

        # filepath = r'C:\20170711\EFWB\32A_E\FAIL\101138-732149843.txt'
        press_count = 0
        # with open(filepath, "r") as ins:
            # line = readfile_lh.readline()
            # for line in ins:
        for line in filecontect:
            if line.find('Button Calibration (Released) FA:') > -1:
                # print(line)
                re_h = re.match('.*FA: (.*),', line)
                if re_h:
                    temp_str = re_h.group(1)
                    temp_str = re.sub(' ', '', temp_str)
                    temp_str = int(temp_str)
                    # print(temp_str)
                    release_value_ll.append(temp_str)
                    release_log_ll.append(line)

                    # print(temp_str)
                else:
                    print('error1')
                    continue
            if line.find('Button Calibration (Pressed) FA:') > -1:
                # print(line)
                re_h = re.match('.*FA: (.*)', line)
                if re_h:
                    press_count = press_count + 1
                    temp_str = re_h.group(1)
                    temp_str = re.sub(' ', '', temp_str)
                    # print(temp_str)
                    # if isinstance(temp_str, int):
                    try:
                        press_value_ls = int(temp_str)
                    except:
                        continue
                    temp_str = release_value_ll[-1]
                    release_press_ll.append(temp_str)
                    release_press_ll.append(press_value_ls)
                    release_press_value_odic[press_count] = release_press_ll
                    release_press_ll = []

                    temp_str = release_log_ll[-1]
                    release_press_log_ll.append(temp_str)
                    release_press_log_ll.append(line)
                    temp_release_press_log_odic[press_count] = release_press_log_ll
                    release_press_log_ll = []

                    # print(temp_release_press_log_odic[press_count])
                else:
                    print('error2')
                    continue

        for key, value in release_press_value_odic.items():
            # print(key)
            # print(value)
            count = 0
            release_value = 0
            press_value = 0
            for i in value:
                count += 1
                if count == 1:
                    release_value = i
                else:
                    press_value = i

            if release_value and press_value:
                if (release_value - press_value) <= 60:
                    is_button_clib_fail_type = 0
                    log_count = 0
                    release_log = ''
                    press_log = ''
                    for i in temp_release_press_log_odic[key]:
                        log_count += 1
                        if log_count == 1:
                            release_log = i
                        else:
                            press_log = i
                    if release_log and press_log:
                        release_press_log_odic[release_log] = press_log
                elif (release_value - press_value) <= 70 and (release_value - press_value) > 60:
                    is_button_clib_fail_type = 1
                    log_count = 0
                    release_log = ''
                    press_log = ''
                    for i in temp_release_press_log_odic[key]:
                        log_count += 1
                        if log_count == 1:
                            release_log = i
                        else:
                            press_log = i
                    if release_log and press_log:
                        release_press_log_odic[release_log] = press_log


        # for k, v in release_press_log_odic.items():
        #     print(k)
        #     print(v)
        # print(is_button_clib_fail)
        return is_button_clib_fail_type, release_press_log_odic

    def check_motion_clib_fail(self, filecontect):
        over_range_count = 0
        less_range_count = 0
        motion_clib_error_type = 0
        motion_clib_error_log_ll = []
        # with open(filepath, "r") as ins:
        for line in filecontect:
            if line.find('Motion Calibration Count: ') > -1:
                # print(line)
                re_h = re.match('.*SPEC\((.*)\)', line)
                range_value = re_h.group(1).split('~')
                clib_minimum = range_value[0]
                clib_maximum = range_value[1]
                re_h = re.match('.*Motion Calibration Count: (\d+),', line)
                if re_h:
                    clib_value = re_h.group(1)
                    clib_value = re.sub(' ', '', clib_value)
                    clib_maximum = int(clib_maximum)
                    clib_minimum = int(clib_minimum)
                    clib_value = int(clib_value)
                    # print(clib_maximum, clib_minimum, clib_value)
                    if clib_value > clib_maximum:
                        over_range_count += 1
                        motion_clib_error_log_ll.append(line)
                    if clib_value < clib_minimum:
                        less_range_count += 1
                        motion_clib_error_log_ll.append(line)

        if less_range_count >= 10:
            motion_clib_error_type = 2
        elif (over_range_count + less_range_count) >= 10:
            motion_clib_error_type = 1

        # print(over_range_count)
        # print(less_range_count)

        return motion_clib_error_type, motion_clib_error_log_ll

    def check_press_twice_fail(self, filecontect):
        twice_fail_log_ll = []
        is_twice_fail = False
        # with open(filepath, "r") as ins:
        for line in filecontect:
            if line.find('Press Crown Twice Again') > -1:
                twice_fail_log_ll.append(line)
            elif line.find('Button Test Result Fail') > -1:
                twice_fail_log_ll.append(line)

        if len(twice_fail_log_ll) == 2:
            is_twice_fail = True

        return is_twice_fail, twice_fail_log_ll

    def check_no_release_clib_fa_fail(self, filecontect):
        no_release_clib_fa_log_ll = []
        is_no_release_clib_fa_fail = False
        # with open(filepath, "r") as ins:
        for line in filecontect:
            if line.find('Button Calibration (Released) Fail') > -1:
                no_release_clib_fa_log_ll.append(line)
            elif line.find('Button Test Result Fail') > -1:
                no_release_clib_fa_log_ll.append(line)

        if len(no_release_clib_fa_log_ll) == 2:
            is_no_release_clib_fa_fail = True

        return is_no_release_clib_fa_fail, no_release_clib_fa_log_ll

    def read_need_txt_contect(self, file_path):
        temp_txt_file = 'temp.txt'
        try:
            readfile_txt_lh = open(file_path, 'r')
            txt_content = str(readfile_txt_lh.read())
            readfile_txt_lh.close()
        except:
            print('error')
            return None
        # print(txt_content)
        list_patten = re.compile('.*(ErrMsg ====>.*)', re.S)
        search_result = re.search(list_patten, txt_content)
        if search_result:
            search_content = search_result.group(1)
        else:
            search_content = txt_content
        # print(search_result.group(1))

        # try:
        #     writefile_tmptxt_lh = open(temp_txt_file, 'w')
        #     writefile_tmptxt_lh.write(search_content)
        #     writefile_tmptxt_lh.close()
        # except:
        #     return None

        # return temp_txt_file
        return search_content

    def find_txtfile_error_and_store_to_dic(self, key, filepath):
        sn_and_error_odic = OrderedDict()
        button_clib_fail_logs_odic = OrderedDict()
        error_ll = []
        temp_str = ''
        find_error = False

        # get last test log in log file
        need_txt_contect = self.read_need_txt_contect(filepath)
        if not need_txt_contect:
            self.setlog("讀取檔錯誤:" + filepath, 'error')
            print(filepath)
            return None

        need_txt_contect_lt = tuple(need_txt_contect.splitlines())

        # check button calibration fail
        is_button_clib_fail_type, button_clib_fail_logs_odic = self.check_button_clib_fail(need_txt_contect_lt)
        if is_button_clib_fail_type >= 0:
            for k, v in button_clib_fail_logs_odic.items():
                temp_str = ('%s,%s,' % (k, v))
            if is_button_clib_fail_type == 0:
                temp_str += r'Button press Released FA-Pressed FA <60'
            if is_button_clib_fail_type == 1:
                temp_str += r'Button press Released FA-Pressed FA >70 and <60'
            error_ll.append(temp_str)

        # check motion calibration fail then 10 times
        temp_str = ''
        motion_clib_failtype, motion_clib_faillog_ll = self.check_motion_clib_fail(need_txt_contect_lt)
        if motion_clib_failtype > 0:
            if motion_clib_failtype == 1:
                for i in motion_clib_faillog_ll:
                    temp_str += ('%s,' % i)
                temp_str += ('Motion Calibration Count out of spec')
            if motion_clib_failtype == 2:
                for i in motion_clib_faillog_ll:
                    temp_str += ('%s,' % i)
                temp_str += ('Motion Calibration Count < SPEC in 10 times')
            error_ll.append(temp_str)

        # check press twice case
        temp_str = ''
        is_twice_fail, twice_fail_log_ll = self.check_press_twice_fail(need_txt_contect_lt)
        if is_twice_fail:
            for i in twice_fail_log_ll:
                temp_str += ('%s,' % i)
            error_ll.append(temp_str)

        # check no Button Calibration (Released) FA fail:
        temp_str = ''
        is_no_clib_fa, no_clib_fa_log_ll = self.check_no_release_clib_fa_fail(need_txt_contect_lt)
        if is_no_clib_fa:
            for i in twice_fail_log_ll:
                temp_str += ('%s,' % i)
            temp_str += r'Empty value of Button Calibration Released FA'
            error_ll.append(temp_str)

        # with open(temp_txt_path, "r") as ins:
            # error_ll.clear()
        for line in need_txt_contect_lt:
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
                    count_maximum = int(count_maximum)
                    count_minimum = int(count_minimum)
                    count_value = int(count_value)
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
                    # print(key)
                    fa_maximum = int(fa_maximum)
                    fa_minimum = int(fa_minimum)
                    for fa_value in fa_value_ll:
                        if fa_value:
                            fa_value = int(fa_value)
                        else:
                            continue
                        if (fa_value > fa_maximum) or (fa_value < fa_minimum):
                            fa_fail = True

            # check Button Calibration FA not in SPEC fail
            temp_str = ''
            if line.find('Button Calibration (Released) FA: ') > -1:
                re_h = re.match('.*FA: (.*)', line)
                if re_h:
                    temp_value = re_h.group(1)
                    temp_value = re.sub(' ', '', temp_value)
                    if temp_value:
                        if not temp_value.find('SPEC') > -1:
                            try:
                                button_clib_fa = int(temp_value)
                                if button_clib_fa < 135:
                                    temp_str = '%s,%s' % (line, r'Motion Calibration Count less of spec')
                                    error_ll.append(temp_str)
                            except:
                                pass

            # check Shutter and FA 0 fail
            btn_release_fail = False
            if line.find(', FA: ') > -1:
                re_h = re.match('.*FA: (.*)', line)
                if re_h:
                    temp_value = re_h.group(1)
                    temp_value = re.sub(' ', '', temp_value)
                    if temp_value:
                        try:
                            btn_release_value = int(temp_value)
                        except:
                            btn_release_value = 0

                        if btn_release_value == 0:
                            btn_release_fail = True

                    else:
                        btn_release_fail = True

                    if btn_release_fail:
                        temp_str = '%s,%s' % (line, r'Button Release Fail - FA: 0')
                        error_ll.append(temp_str)


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

    def get_file_and_fixture_dic(self, sn_path_dic):
        sn_and_fixture_odic = OrderedDict()

        for sn in sn_path_dic.keys():
            try:
                # readfile_html_lh = open('test.html', 'r')
                readfile_lh = open(sn_path_dic[sn], 'r')
                # file_content = str(readfile_lh.read())
            except:
                self.setlog("Error! 讀取 " + sn_path_dic[sn] + " 檔發生錯誤! ", 'error')
                return None

            while True:
                content_line = readfile_lh.readline()
                re_h = re.match('.*;(SIMPLE1_[\d]?);.*', content_line)
                if re_h:
                    sn_and_fixture_odic[sn] = re_h.group(1)
                    break
                if content_line == '':
                    self.setlog("Error! 找不到 " + sn_path_dic[sn] + " 的線別!! ", 'error')
                    return None
            readfile_lh.close()

        # for k, v in sn_and_fixture_odic.items():
        #     print(k)
        #     print(v)

        return sn_and_fixture_odic

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
                count_maximum = int(count_maximum)
                count_minimum = int(count_minimum)
                count_value = int(count_value)
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
                fa_maximum = int(fa_maximum)
                fa_minimum = int(fa_minimum)
                for fa_value in fa_value_ll:
                    if fa_value:
                        fa_value = int(fa_value)
                    else:
                        continue
                    if (fa_value > fa_maximum) or (fa_value < fa_minimum):
                        input_str = input_str.replace('%s%s' % (fa_value, ','), '%s%s%s%s' % (r'<font color="red">', fa_value, ',', r'</font>'))
            return input_str

        else:
            return one_line_fail_text


    def check_fixture_idle(self, input_str):
        is_fixture_idle = False
        if input_str.find('Motion FA: ') > -1:
            re_h = re.match('.*Motion FA: (.*) SPEC\(', input_str)
            if re_h:
                temp_str = re.sub(' ', '', re_h.group(1))
                temp_str = re.sub(',,', '', temp_str)
                fa_value_ll = temp_str.split(',')
                if len(fa_value_ll) == 1:
                    for fa in fa_value_ll:
                        if fa and fa <= 0:
                            is_fixture_idle = True

            else:
                is_fixture_idle = True

        if input_str.find('Motion Count: ') > -1:
            re_h = re.match('.*Motion Count: (\d+),', input_str)
            if re_h:
                count_value = re_h.group(1)
                count_value = re.sub(' ', '', count_value)
                count_value = int(count_value)
                if count_value <= 0:
                    is_fixture_idle = True
            else:
                is_fixture_idle = True

        return is_fixture_idle



    def analysis_log_and_gen_report(self):
        sn_and_csverror_odic = OrderedDict()
        local_csv_path_odic = OrderedDict()
        txt_sn_and_filepath_odic = OrderedDict()
        sn_and_txterror_odic = OrderedDict()
        sn_and_fixture_odic = OrderedDict()
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
        if self.user_input_txtpath.find(';') > -1:
            self.user_input_txtpath = re.sub(r"\n", "", self.user_input_txtpath)
            sn_user_ll = self.user_input_txtpath.split(';')
        else:
            sn_user_ll = self.user_input_txtpath.split()

        # remove duplicate sn number in user input
        temp_ll = sn_user_ll
        sn_user_ll = sorted(set(temp_ll), key=temp_ll.index)
        sn_user_ll = tuple(sn_user_ll)
        # sn_user_ll = list(sn_user_ll)
        # sn_user_ll.sort()

        for input_sn in sn_user_ll:
            for full_path in tuple(all_txt_path):
                full_file_path = os.path.split(full_path)
                file_name_and_ext = os.path.splitext(full_file_path[1])
                file_name = file_name_and_ext[0]
                file_sn = re.sub('.*-', '', file_name)
                if input_sn == file_sn:
                    sn_ll.append(file_name)

        txt_sn_and_filepath_odic = self.get_file_list_store_to_dic(all_txt_path)
        sn_and_fixture_odic = self.get_file_and_fixture_dic(txt_sn_and_filepath_odic)

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

        self.setlog("產生html報表", 'info')
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
            temp_str = ''

            if (sn in sn_and_txterror_odic) or (sn in sn_and_csverror_odic):  # if sn have error that in txt file
                # print('-----%s-----' % k)
                # write sn
                if html_content.find('{{sn}}') == -1:
                    self.setlog("Error! 寫入html發生錯誤, 找不到{{sn}}標籤! ", 'error')
                    return
                html_content = html_content.replace('{{sn}}', sn)

                # write fixture
                temp_str = sn_and_fixture_odic[sn]
                if temp_str:
                    html_content = html_content.replace('{{fixture}}', temp_str)
                else:
                    html_content = html_content.replace('{{fixture}}', 'Fixture not find')

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
        html_content = html_content.replace('{{fixture}}', '')
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
        self.setlog("已產生fail log報表, 請打開 " + html_full_path, 'info')
        # self.store_file_to_folder(html_full_path, '%s%s%s' % (log_foldername, '\\', 'result.xls'))

        # todo: generate excel format form================================================================
        # copy temp.html to log folder
        self.store_file_to_folder('temp_form.xls', log_foldername)
        # rename temp.html to result.html
        os.rename(('%s\\temp_form.xls' % log_foldername), ('%s\\form.xls' % log_foldername))
        excel_full_path = ('%s\\form.xls' % log_foldername)

        self.setlog("產生excel報表", 'info')
        # write to html
        try:
            readfile_excel_lh = open(excel_full_path, 'r')
            excel_content = str(readfile_excel_lh.read())
            readfile_excel_lh.close()
        except Exception as ex:
            self.setlog("Error! 讀取excel檔發生錯誤! "
                        "請確認" + TITLE + "目錄有temp_form.xls且內容正確", 'error')
            return

        for sn in sn_ll:  # read all sn that in user input sn list
            sn_has_txt_error = False
            temp_str = ''

            if (sn in sn_and_txterror_odic) or (sn in sn_and_csverror_odic):  # if sn have error that in txt file

                excel_content = excel_content.replace('{{next_item}}', form_template_need_repeat)

                # write sn
                if excel_content.find('{{sn}}') == -1:
                    self.setlog("Error! 寫入excel發生錯誤, 找不到{{sn}}標籤! ", 'error')
                    return
                excel_content = excel_content.replace('{{sn}}', sn)

                # write fixture
                temp_str = sn_and_fixture_odic[sn]
                if temp_str:
                    excel_content = excel_content.replace('{{fixture}}', temp_str)
                else:
                    excel_content = excel_content.replace('{{fixture}}', 'Fixture not find')

                if sn in csv_sn_and_filepath_odic.keys():

                    # write csv file path to url
                    excel_content = excel_content.replace('{{csv_url}}', csv_sn_and_filepath_odic[sn])
                    # write csv file name
                    excel_content = excel_content.replace('{{csv_filename}}', ('%s%s' % (sn, '_calib.csv')))

                    if sn in sn_and_csverror_odic.keys():
                        # write csv error log to html
                        temp_error_log = ''
                        for fail_list in sn_and_csverror_odic[sn]:
                            if isinstance(fail_list, list):
                                for i in fail_list:
                                    temp_error_log += '%s%s' % (i, '<br>')
                            else:
                                temp_error_log += '%s%s' % (fail_list, '<br>')
                        excel_content = excel_content.replace('{{csv_fail_log}}', temp_error_log)
                    else:
                        excel_content = excel_content.replace('{{csv_fail_log}}', 'No error')
                        # self.setlog("file: " + sn + '沒有csv fail log, 但有txt fail log', 'info')
                else:
                    # write csv file path to url
                    excel_content = excel_content.replace('{{csv_url}}', 'No file')
                    # write csv file name
                    excel_content = excel_content.replace('{{csv_filename}}', 'No file')
                    excel_content = excel_content.replace('{{csv_fail_log}}', 'No file')
                    # self.setlog("file: " + sn + ' 有txt記錄檔但沒有csv檔', 'info')

                # write txt all file path to html
                temp_txt_filepath = ''
                if sn in sn_and_txterror_odic.keys():
                    # write txt file path to url
                    excel_content = excel_content.replace('{{txt_url}}', txt_sn_and_filepath_odic[sn])
                    # write txt file name
                    excel_content = excel_content.replace('{{txt_file_name}}', ('%s%s' % (sn, '.txt')))

                    temp_error_log = ''
                    temp_error_type_ll = []
                    temp_error_type_ls = ''
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

                        # write comment
                        # case 1: button fail and no csv file
                        is_usb_disconnect = False
                        if (fail_list.find('Button Test Result Fail') > -1) and (sn not in csv_sn_and_filepath_odic.keys()):
                            is_usb_disconnect = True
                        elif (fail_list.find('Button Test Enable Fail') > -1) and (sn not in csv_sn_and_filepath_odic.keys()):
                            is_usb_disconnect = True
                        elif (fail_list.find('Record Value Fail') > -1) and (sn not in csv_sn_and_filepath_odic.keys()):
                            is_usb_disconnect = True
                        elif (fail_list.find('Motion Calibration Count less of spec') > -1) and (sn not in csv_sn_and_filepath_odic.keys()):
                            is_usb_disconnect = True
                        elif (fail_list.find('Press Crown Twice Again') > -1) and (sn not in csv_sn_and_filepath_odic.keys()):
                            is_usb_disconnect = True
                        elif fail_list.find('Empty value of Button Calibration Released FA') > -1:
                            if (sn not in sn_and_csverror_odic.keys()) or (sn not in csv_sn_and_filepath_odic.keys()):
                                is_usb_disconnect = True

                        if is_usb_disconnect:
                            temp_str = 'USB Disconnect'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 2: device offline
                        if fail_list.find('device offline') > -1:
                            temp_str = 'device offline'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')
                        # case 3: no devices found
                        if fail_list.find('no devices found') > -1:
                            temp_str = 'no devices found'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 4: Motion FA value is out of SPEC
                        if fail_list.find('Motion FA: ') > -1:
                            re_h = re.match('.*SPEC\((.*)\)', fail_list)
                            range_value = re_h.group(1).split('~')
                            fa_minimum = range_value[0]
                            fa_maximum = range_value[1]
                            re_h = re.match('.*Motion FA: (.*) SPEC\(', fail_list)
                            if re_h:
                                temp_str = re.sub(' ', '', re_h.group(1))
                                temp_str = re.sub(',,', '', temp_str)
                                fa_value_ll = temp_str.split(',')
                                fa_maximum = int(fa_maximum)
                                fa_minimum = int(fa_minimum)
                                for fa_value in fa_value_ll:
                                    if fa_value:
                                        fa_value = int(fa_value)
                                    else:
                                        continue
                                    if (fa_value > fa_maximum) or (fa_value < fa_minimum):
                                        temp_str = '治具轉動異常: Rotation(Motion) FA out of spec'
                                        if not temp_error_type_ls.find(temp_str) > -1:
                                            temp_error_type_ls += '%s%s' % (temp_str, '<br>')
                                            break
                        # case 5: 治具轉動異常
                        if (fail_list.find("Roating On Fail") > -1) or (fail_list.find('Roating Down Fail') > -1):
                            temp_str = '治具轉動異常: Roating ON/Down Fail'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 6: 治具按壓異常
                        if fail_list.find(r'Button press Released FA-Pressed FA <60') > -1:
                            temp_str = '治具按壓異常: Button press Released FA-Pressed FA <60'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')
                        # case 7: fixture idle
                        if self.check_fixture_idle(fail_list):
                            temp_str = '治具空轉'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')
                        # case 8: Motion Count out of spec
                        if fail_list.find('Motion Count: ') > -1:
                            # print(input_str)
                            re_h = re.match('.*SPEC\((.*)\)', fail_list)
                            range_value = re_h.group(1).split('~')
                            count_minimum = range_value[0]
                            count_maximum = range_value[1]
                            re_h = re.match('.*Motion Count: (\d+),', fail_list)
                            if re_h:
                                count_value = re_h.group(1)
                                count_value = re.sub(' ', '', count_value)
                                count_value = int(count_value)
                                if count_value > 300:
                                    temp_str = 'Motion Count out of spec, will fix on sw270 version'
                                    if not temp_error_type_ls.find(temp_str) > -1:
                                        temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 9: Motion Calibration Count out of spec
                        # case 12: 組裝(Crown sensor 偏移)
                        if fail_list.find('Motion Calibration Count out of spec') > -1:
                            temp_str = 'Motion Calibration Count out of spec: 240+-30 to fix'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')
                        if fail_list.find('Motion Calibration Count < SPEC in 10 times') > -1:
                            temp_str = '組裝(Crown sensor 偏移'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 10: 暫時不用管
                        # case 11: 組裝(Board to Board 沒扣)
                        if (fail_list.find('Button Calibration (Released) Fail') > -1) and (fail_list.find('Button Calibration (Released) FA: T9127]Open CSV file success') > -1):
                            if 'Check button and motion status FAIL' in sn_and_csverror_odic[sn]:
                                temp_str = '組裝 - Board to Board 沒扣'
                                if not temp_error_type_ls.find(temp_str) > -1:
                                    temp_error_type_ls += '%s%s' % (temp_str, '<br>')
                        elif (fail_list.find('Button Calibration (Released) Fail') > -1) and (sn not in csv_sn_and_filepath_odic.keys()):
                            temp_str = '組裝 - Board to Board 沒扣'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 13: 組裝 - 亮度不足(Crown sensor 偏移)
                        if fail_list.find('Button Calibration (Released) Fail') > -1:
                            if sn in sn_and_csverror_odic.keys():
                                for csv_error in sn_and_csverror_odic[sn]:
                                    if csv_error.find('FAIL! SHUTTER_C already reach min level: 0') > -1:
                                        temp_str = '組裝 - 亮度不足(Crown sensor 偏移)'
                                        if not temp_error_type_ls.find(temp_str) > -1:
                                            temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 14: 治具轉動異常M Rotation(Motion) Count out of spec, Count +-30
                        if fail_list.find('Motion Count:') > -1:
                            re_h = re.match('.*SPEC\((.*)\)', fail_list)
                            range_value = re_h.group(1).split('~')
                            count_minimum = range_value[0]
                            count_maximum = range_value[1]
                            re_h = re.match('.*Motion Count: (\d+),', fail_list)
                            if re_h:
                                count_value = re_h.group(1)
                                count_value = re.sub(' ', '', count_value)
                                count_value = int(count_value)
                                count_minimum = int(count_minimum)
                                if count_value < count_minimum:
                                    if (sn not in sn_and_csverror_odic.keys()) or (sn not in csv_sn_and_filepath_odic.keys()):
                                        temp_str = 'SW 280 fixed'
                                        if not temp_error_type_ls.find(temp_str) > -1:
                                            temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 15: SW 270 fixed
                        if fail_list.find('Motion Calibration Count less of spec') > -1:
                            if (sn in sn_and_csverror_odic.keys()):
                                for csv_fail in sn_and_csverror_odic[sn]:
                                    if csv_fail.find('FAIL! Can\'t calibration FA with SHUTTER_C/F') > -1:
                                        temp_str = 'SW 270 fixed FA out of SPEC.'
                                        if not temp_error_type_ls.find(temp_str) > -1:
                                            temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 16: Button press Released FA-Pressed FA >70 and <60
                        if fail_list.find(r'Button press Released FA-Pressed FA >70 and <60') > -1:
                            temp_str = '治具按壓異常 - Button press Released FA-Pressed FA<70 - SW 280 fixed'
                            if not temp_error_type_ls.find(temp_str) > -1:
                                temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # case 17: fix 280 sw
                        if fail_list.find(r'Button Release Fail - FA: 0') > -1:
                            if (sn in sn_and_csverror_odic.keys()):
                                for csv_fail in sn_and_csverror_odic[sn]:
                                    if csv_fail.find('FAIL! SHUTTER_C already reach max level') > -1:
                                        temp_str = 'SW 280 imporve'
                                        if not temp_error_type_ls.find(temp_str) > -1:
                                            temp_error_type_ls += '%s%s' % (temp_str, '<br>')

                        # temp_error_log += '%s%s' % ('*******', '<br>')
                    excel_content = excel_content.replace('{{txt_fail_log}}', temp_error_log)
                    excel_content = excel_content.replace('{{comment}}', temp_error_type_ls)

                else:
                    # write txt file path to url
                    excel_content = excel_content.replace('{{txt_url}}', txt_sn_and_filepath_odic[sn])
                    # write txt file name
                    excel_content = excel_content.replace('{{txt_file_name}}', ('%s%s' % (sn, '.txt')))
                    excel_content = excel_content.replace('{{txt_fail_log}}', 'No error')

                # excel_content = excel_content.replace('{{next_item}}', form_template_need_repeat)

        excel_content = excel_content.replace('{{next_item}}', '')
        excel_content = excel_content.replace('{{sn}}', '')
        excel_content = excel_content.replace('{{fixture}}', '')
        excel_content = excel_content.replace('{{csv_filename}}', '')
        excel_content = excel_content.replace('{{csv_fail_log}}', '')
        excel_content = excel_content.replace('{{txt_file_name}}', '')
        excel_content = excel_content.replace('{{txt_fail_log}}', '')
        # excel_content = excel_content.replace('{{comment}}', '')

        try:
            # writhfile_html_lh = open('test.html', 'w')
            writhfile_excel_lh = open(excel_full_path, 'w')
            writhfile_excel_lh.write(excel_content)
            writhfile_excel_lh.close()

        except Exception as ex:
            self.setlog("Error! 寫入excel檔發生錯誤! "
                        "請確認" + TITLE + "目錄有temp_form.xls且內容正確", 'error')
            return

        # self.store_file_to_folder(excel_full_path, '%s%s%s' % (log_foldername, '\\', 'result.xls'))
        self.setlog("已產生excel報表, 請打開 " + excel_full_path, 'info')


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
