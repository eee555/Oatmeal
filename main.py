import json, math, re, xlrd
import openpyxl
import os


# pip install xlrd==1.2.0

# 根据excel文件生成data.ts文件
# excel文件是从pms3.0系统上导下来的
# 生成名为data.ts的数据文件
# 第一个sheet是220的断路器，名字任意
# 第而个sheet是500的断路器，名字任意

import tkinter as tk
from tkinter import filedialog

class ExcelReader:
    def __init__(self, file_path, sheet=None):
        self.file_path = file_path
        self.sheet_name = sheet
        self.workbook = None
        self.sheet = None
        self._open_workbook()

    def _open_workbook(self):
        if self.file_path.endswith('.xlsx'):
            # 处理 .xlsx 文件
            self.workbook = openpyxl.load_workbook(self.file_path)
            if self.sheet_name:
                self.sheet = self.workbook[self.sheet_name]
            else:
                self.sheet = self.workbook.active
        elif self.file_path.endswith('.xls'):
            # 处理 .xls 文件
            self.workbook = xlrd.open_workbook(self.file_path)
            if self.sheet_name:
                self.sheet = self.workbook.sheet_by_name(self.sheet_name)
            else:
                self.sheet = self.workbook.sheet_by_index(0)
        else:
            raise ValueError("不支持的文件格式，仅支持 .xlsx 和 .xls 文件。")

    def __iter__(self):
        if self.sheet is None:
            return
        if self.file_path.endswith('.xlsx'):
            # 遍历 .xlsx 文件的每一行
            for row in self.sheet.iter_rows(min_row=2, values_only=True):
                if any(row):
                    yield row
        elif self.file_path.endswith('.xls'):
            # 遍历 .xls 文件的每一行
            for row_index in range(1, self.sheet.nrows):
                row = self.sheet.row_values(row_index)
                # 检查行是否为空
                if any(row):
                    yield row

    def close(self):
        if self.file_path.endswith('.xlsx'):
            self.workbook.close()
            

def 线路数据整理(xlsx_file_name_1="断路器20241126（220千伏开关及500千伏开关）.xls",
           xlsx_file_name_2="500中心保护类型.xlsx", file_output=True):

    # xlsx_file_name_1 = "断路器20241126（220千伏开关及500千伏开关）.xls"
    # xlsx_file_name_2 = "500中心保护类型.xlsx"
    
    xls_sheet_220 = ExcelReader(xlsx_file_name_1, "220千伏开关")
    xls_sheet_500 = ExcelReader(xlsx_file_name_1, "500千伏开关")
    
    
    xls_sheet_500_lines = ExcelReader(xlsx_file_name_2, "线路")
    xls_sheet_500_transformers = ExcelReader(xlsx_file_name_2, "主变")
    xls_sheet_500_buses = ExcelReader(xlsx_file_name_2, "母线")
    
    
    # 清洗线路名称
    def clean_line_name(ol_line):
        line = ol_line
        line = line.replace(" ", "").replace("220kV", "").replace("开关", "").\
            replace("断路器", "").replace("220KV", "").replace("220kv", "").\
                replace("500KV", "").replace("500kV", "").replace("500kv", "")
        if line[2] == "线":
            line = line[:2] + line[3:]
        if len(line) >= 6 and line[5] == "线":
            line = line[:5] + line[6:]
        if len(line) != 6 and len(line) != 9 and "分段" not in line and "主变" not in line and\
            "母联" not in line and "内桥" not in line and "旁路" not in line:
                # 清洗不干净的打印出来
            print(ol_line)
        return line
    
    # 交换前两个字符
    def swap_12(line):
        return line[1] + line[0] + line[2:]
    
    # 把nan换成“”
    def n2n(a):
        if isinstance(a, str):
            return a
        if a==None:
            return ""
        return "" if math.isnan(a) else str(a)
        
    
    # 整理出500变电站、开关
    bdz_500 = set()
    breaker_500 = []
    for row in xls_sheet_500:
        bdz_500 = bdz_500 | { row[3] }
        breaker_500.append({"name": clean_line_name(row[0]), "bdz": row[3]})
      
    data1 = []  
    # 整理出220变电站、开关
    bdz_220 = set()
    for row in xls_sheet_220:
        bdz_220 = bdz_220 | { row[3] }
        # 备用线路故障忽略
        if "备用" in row[0]:
            continue
        data1.append((clean_line_name(row[0]), row[3]))
    bdz_220 = bdz_220 - bdz_500
        
    
    # 整理所有500开关
    # breaker_500 = []
    # for i in range(1, xls_sheet_500.nrows):
    #     rowdate = xls_sheet_500.row_values(i)#i行的list
    #     breaker_500.append({"name": clean_line_name(rowdate[0]), "bdz": rowdate[3]})
    
        
    line_dict_220 = {}
    line_list_220 = []
    for i in data1:
        if "分段" in i[0] or "主变" in i[0] or "母联" in i[0] or\
            "内桥" in i[0] or "旁路" in i[0]:
            i0 = i[0].replace("III", "Ⅲ").replace("IV", "Ⅳ").\
                replace("ⅠⅠ", "Ⅱ").replace("ⅠV", "Ⅳ").replace("II", "Ⅱ").\
                    replace("I", "Ⅰ")
            if "分段" in i0 and "段分段" not in i0 and i0[:2] != "分段":
                i0 = i0.replace("分段", "段分段")
            if "母联" in i0 and "段母联" not in i0 and i0[:2] != "母联":
                i0 = i0.replace("母联", "段母联")
            if i0[0] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"] and\
                i0[1] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"]:
                i0 = i0[0] + "、" + i0[1:]
            if i0[0] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"] and\
                i0[2] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"] and i0[1] == '-':
                i0 = i0[0] + "、" + i0[2:]
            i0 = i0.replace("Ⅰ", "I").replace("Ⅱ", "II").\
                replace("Ⅲ", "III").replace("Ⅳ", "IV")
            line_list_220.append((i0, i[1] + " "))
            continue
        if i[0] not in line_dict_220 and swap_12(i[0]) not in line_dict_220:
            if i[0][0] == i[1][0]:
                line_dict_220.update({i[0]: i[1] + " "})
            elif i[0][0] == i[1][1]:
                line_dict_220.update({i[0]: i[1] + " "})
            elif i[0][1] == i[1][0]:
                line_dict_220.update({i[0]: " " + i[1]})
            elif i[0][1] == i[1][1]:
                line_dict_220.update({i[0]: " " + i[1]})   
            else:
                print(i)
        else:
            if i[0] in line_dict_220:
                key = i[0]
            else:
                key = swap_12(i[0])
            if key[0] == i[1][0]:
                if line_dict_220[key][0] != " ":
                    print(i)
                line_dict_220[key] = i[1] + line_dict_220[key]
            elif key[0] == i[1][1]:
                if line_dict_220[key][0] != " ":
                    print(i)
                line_dict_220[key] = i[1] + line_dict_220[key]
            elif key[1] == i[1][0]:
                if line_dict_220[key][-1] != " ":
                    print(i)
                line_dict_220[key] = line_dict_220[key] + i[1]
            elif key[1] == i[1][1]:
                if line_dict_220[key][-1] != " ":
                    print(i)
                line_dict_220[key] = line_dict_220[key] + i[1]
            else:
                print(i)
    for i in range(len(line_list_220)):
        a = line_list_220[i][1].split(" ")
        line_list_220[i] = {"name": line_list_220[i][0], "start": a[0], "end": a[1]}
    for i in line_dict_220.keys():
        a = line_dict_220[i].split(" ")
        line_list_220.append({"name": i, "start": a[0], "end": a[1]})
    # 删去两端都不是220站的线路
    line_list_220 = filter(lambda x: x['start'] in bdz_220 or x['end'] in\
                           bdz_220, line_list_220)
    # 首端不是220站的线路改为首端是220站的线路
    line_list_220 = map(lambda x: x if x['start'] in bdz_220 else\
                        {"name": swap_12(x["name"]), "start": x["end"],
                         "end": x["start"]}, line_list_220)
    line_list_220 = list(line_list_220)
    
        
    line_dict_500 = {}
    line_list_500 = []
    for i in data1:
        if "分段" in i[0] or "主变" in i[0] or "母联" in i[0] or\
            "内桥" in i[0] or "旁路" in i[0]:
            i0 = i[0].replace("III", "Ⅲ").replace("IV", "Ⅳ").\
                replace("ⅠⅠ", "Ⅱ").replace("ⅠV", "Ⅳ").replace("II", "Ⅱ").\
                    replace("I", "Ⅰ")
            if "分段" in i0 and "段分段" not in i0 and i0[:2] != "分段":
                i0 = i0.replace("分段", "段分段")
            if "母联" in i0 and "段母联" not in i0 and i0[:2] != "母联":
                i0 = i0.replace("母联", "段母联")
            if i0[0] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"] and\
                i0[1] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"]:
                i0 = i0[0] + "、" + i0[1:]
            if i0[0] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"] and\
                i0[2] in ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"] and i0[1] == '-':
                i0 = i0[0] + "、" + i0[2:]
            i0 = i0.replace("Ⅰ", "I").replace("Ⅱ", "II").\
                replace("Ⅲ", "III").replace("Ⅳ", "IV")
            line_list_500.append((i0, i[1] + " "))
            continue
        if i[0] not in line_dict_500 and swap_12(i[0]) not in line_dict_500:
            if i[0][0] == i[1][0]:
                line_dict_500.update({i[0]: i[1] + " "})
            elif i[0][0] == i[1][1]:
                line_dict_500.update({i[0]: i[1] + " "})
            elif i[0][1] == i[1][0]:
                line_dict_500.update({i[0]: " " + i[1]})
            elif i[0][1] == i[1][1]:
                line_dict_500.update({i[0]: " " + i[1]})   
            else:
                print(i)
        else:
            if i[0] in line_dict_500:
                key = i[0]
            else:
                key = swap_12(i[0])
            if key[0] == i[1][0]:
                if line_dict_500[key][0] != " ":
                    print(i)
                line_dict_500[key] = i[1] + line_dict_500[key]
            elif key[0] == i[1][1]:
                if line_dict_500[key][0] != " ":
                    print(i)
                line_dict_500[key] = i[1] + line_dict_500[key]
            elif key[1] == i[1][0]:
                if line_dict_500[key][-1] != " ":
                    print(i)
                line_dict_500[key] = line_dict_500[key] + i[1]
            elif key[1] == i[1][1]:
                if line_dict_500[key][-1] != " ":
                    print(i)
                line_dict_500[key] = line_dict_500[key] + i[1]
            else:
                print(i)
    for i in range(len(line_list_500)):
        a = line_list_500[i][1].split(" ")
        line_list_500[i] = {"name": line_list_500[i][0], "start": a[0], "end": a[1]}
    for i in line_dict_500.keys():
        a = line_dict_500[i].split(" ")
        line_list_500.append({"name": i, "start": a[0], "end": a[1]})
    # 删去两端都不是500站的线路
    line_list_500 = filter(lambda x: x['start'] in bdz_500 or x['end'] in\
                           bdz_500, line_list_500)
    # 首端不是500站的线路改为首端是500站的线路
    line_list_500 = map(lambda x: x if x['start'] in bdz_500 else\
                        {"name": swap_12(x["name"]), "start": x["end"],
                         "end": x["start"]}, line_list_500)
    line_list_500 = list(line_list_500)
    
    
    # 整理500线路的开关、哪些线路保护停用
    line_list_500_500 = []
    stop_reclose = []
    for row in xls_sheet_500_lines:
    # for i in range(1, xls_sheet_500_lines.nrows):
    #     row = xls_sheet_500_lines.row_values(i)
    # for index, row in xls_sheet_500_lines.iterrows():
        # print(row[3])
        name = row[3].strip("220").strip("500").strip("kV").strip("千伏").strip("线")
        if row[9] == "停用":
            stop_reclose.append(name)
            stop_reclose.append(swap_12(name))
        if row[6] != "500kV":
            continue
        line_list_500_500.append({"name": name, "start": row[1],
                                  "end": row[2], "side": str(row[4]),
                                  "middle": str(row[5]), "p1": str(row[7]),
                                  "p2": str(row[8])})
    
    # 500站的主变三侧开关、保护型号
    transformers_500 = {}
    for row in xls_sheet_500_transformers:
    # for i in range(1, xls_sheet_500_transformers.nrows):
    #     row = xls_sheet_500_transformers.row_values(i)
    # for index, row in xls_sheet_500_transformers.iterrows():
        # print(row[3])
        if row[1] not in transformers_500:
            transformers_500[row[1]] = {}
            
        transformers_500[row[1]][row[2]] = {"breaker_1": n2n(row[3]), "breaker_2": n2n(row[4]),
                                            "breaker_3": n2n(row[5]) or "", "pro_1": n2n(row[6]),
                                            "pro_2": n2n(row[7]), "pro_3": n2n(row[8])}
    
    
    # 500站的220、500母线上的开关、保护
    buses_500 = {}
    for row in xls_sheet_500_buses:
    # for i in range(1, xls_sheet_500_buses.nrows):
    #     row = xls_sheet_500_buses.row_values(i)
    # for index, row in xls_sheet_500_buses.iterrows():
        bdz = row[1]
        if bdz not in buses_500:
            buses_500[bdz] = {}
        ids = re.findall("[A-Z0-9]+", n2n(row[3])) + re.findall("[A-Z0-9]+", n2n(row[4]))
        # for idx, breaker_id in enumerate(ids):
        #     if breaker_id[0] == "5":
        #         ids[idx] = next(filter(lambda x: x["bdz"]==bdz and\
        #                                breaker_id in x["name"], breaker_500))["name"]
        #     else:
        #         ids[idx] = next(filter(lambda x: x[1]==bdz and\
        #                                breaker_id in x[0], data1))[0]
        buses_500[bdz][row[2]] = {"breakers": ids, "pro_1": n2n(row[5]),
                                  "pro_2": n2n(row[6])}
    
    
    str_breaker_500 = "var breaker_500 = "
    str_breaker_500 += json.dumps(breaker_500, ensure_ascii = False)
    str_line_data_220 = "var data_data_base_220 = "
    str_line_data_220 += json.dumps(line_list_220, ensure_ascii = False)
    # all_bdz_name = [i["start"] for i in line_list_220] + [i["end"] for i in line_list_220]
    str_bdz_name_220 = "var data_data_bdz_220 = "
    str_bdz_name_220 += json.dumps(list(bdz_220), ensure_ascii=False)
    
    str_line_data_500_220 = "var data_base_500_220 = "
    str_line_data_500_220 += json.dumps(line_list_500, ensure_ascii = False)
    str_line_data_500_500 = "var data_base_500_500 = "
    str_line_data_500_500 += json.dumps(line_list_500_500, ensure_ascii = False)
    # all_bdz_name = [i["start"] for i in line_list_500] + [i["end"] for i in line_list_500]
    str_bdz_name_500 = "var data_bdz_500 = "
    str_bdz_name_500 += json.dumps(list(bdz_500), ensure_ascii=False)
    
    str_stop_reclose = "var stop_reclose = "
    str_stop_reclose += json.dumps(stop_reclose, ensure_ascii=False)
    
    
    str_transformers_500 = "var transformers_500 = "
    str_transformers_500 += json.dumps(transformers_500, ensure_ascii=False)
    
    
    str_buses_500 = "var buses_500 = "
    str_buses_500 += json.dumps(buses_500, ensure_ascii=False)
    
    file_content = "// 此文件由“整理线路数据脚本.py”自动生成，请勿手动修改。"
    file_content += "\n"
    file_content += str_breaker_500
    file_content += "\n"
    file_content += str_line_data_220
    file_content += "\n"
    file_content += str_bdz_name_220
    file_content += "\n"
    file_content += str_line_data_500_220
    file_content += "\n"
    file_content += str_line_data_500_500
    file_content += "\n"
    file_content += str_bdz_name_500
    file_content += "\n"
    file_content += str_stop_reclose
    file_content += "\n"
    file_content += str_transformers_500
    file_content += "\n"
    file_content += str_buses_500
    file_content += "\n"
    file_content += "// 此行注释勿删，定位用"
    
    if file_output:
        with open('data.ts', 'w', encoding="utf-8") as f:
            str_breaker_500 = "export const breaker_500: { name: string, bdz: string }[] = "
            str_breaker_500 += json.dumps(breaker_500, ensure_ascii = False)
            str_line_data_220 = "export const data_base_220: { name: string, start: string, end: string }[] = "
            str_line_data_220 += json.dumps(line_list_220, ensure_ascii = False)
            # all_bdz_name = [i["start"] for i in line_list_220] + [i["end"] for i in line_list_220]
            str_bdz_name_220 = "export const data_bdz_220: string[] = "
            str_bdz_name_220 += json.dumps(list(bdz_220), ensure_ascii=False)
            
            str_line_data_500_220 = "export const data_base_500_220: {" +\
                " name: string, start: string, end: string }[] = "
            str_line_data_500_220 += json.dumps(line_list_500, ensure_ascii = False)
            str_line_data_500_500 = "export const data_base_500_500: " +\
                "{ name: string, start: string, end: string, side: string," +\
                    " middle: string, p1: string, p2: string }[] = "
            str_line_data_500_500 += json.dumps(line_list_500_500, ensure_ascii = False)
            # all_bdz_name = [i["start"] for i in line_list_500] + [i["end"] for i in line_list_500]
            str_bdz_name_500 = "export const data_bdz_500: string[] = "
            str_bdz_name_500 += json.dumps(list(bdz_500), ensure_ascii=False)
            
            str_stop_reclose = "export const stop_reclose: string[] = "
            str_stop_reclose += json.dumps(stop_reclose, ensure_ascii=False)
            
            
            str_transformers_500 = "export const transformers_500 = "
            str_transformers_500 += json.dumps(transformers_500, ensure_ascii=False)
            
            
            str_buses_500 = "export const buses_500 = "
            str_buses_500 += json.dumps(buses_500, ensure_ascii=False)
            
            file_content2 = "// 此文件由“整理线路数据脚本.py”自动生成，请勿手动修改。"
            file_content2 += "\n"
            file_content2 += str_breaker_500
            file_content2 += "\n"
            file_content2 += str_line_data_220
            file_content2 += "\n"
            file_content2 += str_bdz_name_220
            file_content2 += "\n"
            file_content2 += str_line_data_500_220
            file_content2 += "\n"
            file_content2 += str_line_data_500_500
            file_content2 += "\n"
            file_content2 += str_bdz_name_500
            file_content2 += "\n"
            file_content2 += str_stop_reclose
            file_content2 += "\n"
            file_content2 += str_transformers_500
            file_content2 += "\n"
            file_content2 += str_buses_500
            file_content2 += "\n"
            file_content2 += "// 此行注释勿删，定位用"
            f.writelines(file_content2)
    return file_content

def replace_text_in_file(file, s, start_delimiter, end_delimiter):
    try:
        # 打开文件并读取内容
        with open(file, 'r', encoding='utf-8') as f:
            content = f.read()

        new_content = ""
        
        # 查找开头定位符
        start_index = content.find(start_delimiter, 0)
        # print("start_index", start_index)
        if start_index == -1:
            raise RuntimeError

        # 将开头定位符之前的内容添加到新内容中
        new_content += content[:start_index]

        # 查找结尾定位符
        end_index = content.find(end_delimiter, start_index + len(start_delimiter))
        # print("end_index", end_index)
        if end_index == -1:
            raise RuntimeError

        # 将指定字符串 s 添加到新内容中
        # print(s)
        new_content += s
        
        new_content += content[end_index + len(end_delimiter):]

        # 将替换后的内容写回文件
        with open(file, 'w', encoding='utf-8') as f:
            f.write(new_content)

        print("替换成功！")
    except FileNotFoundError:
        print(f"文件 {file} 未找到。")
    except Exception as e:
        print(f"发生错误: {e}")



def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def on_confirm():
    file_path1 = entry1.get()
    file_path2 = entry2.get()
    # print(f"选择的第一个文件路径: {file_path1}")
    # print(f"选择的第二个文件路径: {file_path2}")
    search_dirs = ['./generate_fault_message_plugin/dist', './dist', '.']

    # 遍历每个目录进行查找
    for directory in search_dirs:
        for root, dirs, files in os.walk(directory):
            if 'app.js' in files:
                # 计算相对路径
                relative_path = os.path.join(root, 'app.js')
                break
        else:
            continue
        break
    else:
        mainapp.destroy()
        return
    try:
        file_content = 线路数据整理(file_path1, file_path2, file_output=True)
    except:
        mainapp.destroy()
        return
    if file_content:
        # print(relative_path)
        replace_text_in_file(relative_path,
                             file_content, 
                             "// 此文件由“整理线路数据脚本.py”自动生成，请勿手动修改。",
                             "// 此行注释勿删，定位用")
    mainapp.destroy()
    
    

# 创建主窗口
mainapp = tk.Tk()
mainapp.geometry("500x320")
mainapp.title("新一代集控信息汇报工具数据修改工具")

# 创建第一个文件选择框
label1 = tk.Label(mainapp, text="选择断路器Excel文件:")
label1.pack(pady=10)
entry1 = tk.Entry(mainapp, width=50)
entry1.pack(pady=5)
button1 = tk.Button(mainapp, text="选择文件", command=lambda: select_file(entry1))
button1.pack(pady=5)

# 创建第二个文件选择框
label2 = tk.Label(mainapp, text="选择保护Excel文件:")
label2.pack(pady=10)
entry2 = tk.Entry(mainapp, width=50)
entry2.pack(pady=5)
button2 = tk.Button(mainapp, text="选择文件", command=lambda: select_file(entry2))
button2.pack(pady=5)

# 创建确定按钮
confirm_button = tk.Button(mainapp, text="确定", command=on_confirm)
confirm_button.pack(pady=20)


# 线路数据整理()

mainapp.mainloop()

# pyinstaller -F -w main.py




