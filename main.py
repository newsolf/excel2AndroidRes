# This is a sample Python script.
import os
import time

import xlrd as xlrd
import xlsxwriter
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from pathlib import Path

DIR_PATH = "xlsFile"
OUT_DIR_PATH = "outputs"
ENCODING = "UTF-8"


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.


def get_xls_name():
    dir_path = Path(DIR_PATH)
    if 1 - dir_path.exists():
        os.mkdir(DIR_PATH)
        print("mkdir path = %s ,please set file in %s " % dir_path)
        return ""
    if 1 - dir_path.is_dir():
        os.remove(dir_path)
        os.mkdir(DIR_PATH)
        print("mkdir path = %s ,please set file in %s " % DIR_PATH, DIR_PATH)
        return None

    file_list = os.listdir(DIR_PATH)
    if len(file_list) == 0:
        print("please set file in %s " % DIR_PATH)
        return None

    file_name = ""
    for index in range(len(file_list)):
        file = file_list[index]
        if file.endswith("result.xls"):
            print('need delete file %s' % file)
            os.remove(os.path.join(DIR_PATH, file))

        if file.endswith(".xls"):
            file_name = str(file)
            return os.path.join(DIR_PATH, file)


def write_file(file, key, v):
    print('%s  %s\t%s ' % (file.name, key, v))
    file.write("<string name=\"%s\">%s</string>\n" % (key, v))
    file.flush()
    pass


def delete_outputs(dir_path):
    if not os.path.exists(dir_path):
        return False
    if os.path.isfile(dir_path):
        print('delete %s' % dir_path)
        os.remove(dir_path)
        return
    file_list = os.listdir(dir_path)
    for sub in file_list:
        t = os.path.join(dir_path, sub)
        print('delete t %s' % t)
        if os.path.isdir(t):
            delete_outputs(t)
        else:
            print('delete unlink  %s' % t)
            os.unlink(t)
    # os.removedirs(dir_path)


def excel_2_android_res():
    print("excel_2_android_res")
    xls_name = get_xls_name()
    if xls_name == "" or xls_name is None:
        print("no file,break")
        return
    print("file is %s" % xls_name)

    zh_file = open(DIR_PATH + "/strings.xml", "r")
    if zh_file == "" or zh_file is None:
        print("zh_file no file,break")
        return

    delete_outputs(OUT_DIR_PATH)

    workbook = xlrd.open_workbook(xls_name)
    print(workbook.sheet_names()[0])
    worksheet = workbook.sheet_by_index(0)
    print(worksheet)
    rows = worksheet.nrows  # 获取该表总行数
    print(rows)  # 198

    cols = worksheet.ncols  # 获取该表总列数
    print(cols)  # 3

    row_line = worksheet.row_values(1)
    res_file_list = []
    for i in range(len(row_line)):
        res_dir_name = row_line[i]
        if res_dir_name != '':
            print(row_line[i])
            res_dir_path = OUT_DIR_PATH + "/res/values-" + res_dir_name
            res_dir = Path(res_dir_path)
            if 1 - res_dir.exists():
                os.makedirs(res_dir_path)

            res_file_path = res_dir_path + "/strings.xml"
            res_file = open(res_file_path, "w+")
            res_file.writelines("<resources>\n")
            # res_file.write("\n")
            res_file.flush()
            # res_file.close()
            # print(res_file)
            res_file_list.append(res_file)

    count = 0
    for line in zh_file.readlines():
        line = line.strip("\n")
        if line.endswith("</string>"):
            # print(line)
            line_split = line.split("name=\"")[1]
            line_split = line_split.split("\">")
            key = line_split[0]
            # print('key = %s' % key)
            value = line_split[1].split("</string>")[0]
            # print('value = %s' % value)
            for i in range(rows):
                xls_num = worksheet.row_values(i)[0]
                xls_zh = worksheet.row_values(i)[1]
                xls_ru = worksheet.row_values(i)[2]
                if xls_zh == value:
                    # print('key %s,value %s,col %s num %d ru %s' % (key, value, i, xls_num, xls_ru))
                    count += 1
                    for col2 in range(cols):
                        if col2 < 1:
                            continue
                        # print(col2)
                        write_file(res_file_list[col2 - 1], key, worksheet.row_values(i)[col2])

    print('process finished count %d' % count)
    for i in range(len(res_file_list)):
        res_f = res_file_list[i]
        res_f.write("</resources>")
        res_f.flush()
        res_f.close()

    # if(worksheet.row_values(i).)
    pass


if __name__ == '__main__':
    start_time = time.time()
    excel_2_android_res()
    end_time = time.time()
    print('use time %s' % (end_time - start_time))

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
