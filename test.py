import re
from collections import OrderedDict

# PATH = r'C:\1110710\EFWB\20170709\32A_E\FAIL\043834-732050032.txt'
PATH = r'C:\20170711\EFWB\32A_E\FAIL\101138-732149843.txt'

def read_need_txt_contect(file_path):
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
    print(search_result.group(1))


read_need_txt_contect(PATH)

def check_button_clib_fail(filepath):
    release_press_ll = []
    release_value_ll = []
    release_log_ll = []
    release_press_log_ll = []
    release_press_value_odic = OrderedDict()
    temp_release_press_log_odic = OrderedDict()
    release_press_log_odic = OrderedDict()
    is_button_clib_fail = False

    press_count = 0
    with open(filepath, "r") as ins:
        # line = readfile_lh.readline()
        # for line in ins:
        for line in ins:
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
            if line.find('Button Calibration (Pressed) FA:') > -1:
                # print(line)
                re_h = re.match('.*FA: (.*)', line)
                if re_h:
                    press_count = press_count + 1
                    temp_str = re_h.group(1)
                    temp_str = re.sub(' ', '', temp_str)
                    # print(temp_str)
                    press_value_ls = int(temp_str)
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
            if (release_value - press_value) < 70:
                is_button_clib_fail = True
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
    return is_button_clib_fail, release_press_log_odic

def test():
    error_ll = []
    with open(PATH, "r") as ins:
        error_ll.clear()
        fa_value_ll = []

        # is_button_clib_fail, button_clib_fail_log_odic = check_button_clib_fail(PATH)


        for line in ins:
            count_fail = False
            fa_fail = False
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
                    # print(count_value)
                    # print(count_maximum, count_minimum, count_value)
                    if (count_value > count_maximum) or (count_value < count_minimum):
                        # print(count_value)
                        count_fail = True
                        # print('count out of range')
            if line.find('Motion FA: ') > -1:
                # print(line)
                re_h = re.match('.*SPEC\((.*)\)', line)
                range_value = re_h.group(1).split('~')
                fa_minimum = range_value[0]
                fa_maximum = range_value[1]
                re_h = re.match('.*Motion FA: (.*) SPEC\(', line)
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
                            fa_fail = True
                            # print(fa_maximum, fa_minimum, fa_value)

                # if fa_fail:
                #     print(fa_maximum, fa_minimum, fa_value)

            # if line.find('Motion FA: ') > -1:
            #     print(line)
    # if file_name_and_ext[1] == r'.csv':
    #     re_h = re.match(r'(.+)_calib', file_name_and_ext[0])
    #     sn_and_filepath_odic[re_h.group(1)] = file