import tkinter
import openpyxl
from openpyxl.styles import Font, Alignment
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

window = tkinter.Tk()
window.title("Отчеты")
window.geometry("400x200")
window.config(background="#F5F5DC")

def pretense():
    ftypes = [('txt файлы', '*.txt'), ('Все файлы', '*')]
    file_path = filedialog.askopenfilename(filetypes=ftypes)
    if file_path == "":
        messagebox.showinfo("Внимание", "Файл не выбран")
    elif file_path != "" and file_path[-3:] != "txt":
        messagebox.showwarning("Внимание!", "Выберите файл формата txt")
    elif file_path != "" and file_path[-3:] == "txt":
        list_tr_f = ['FEE COLL-M2M', 'FEE COLL-MBG', 'FEE COLL-RET', 'FEE COLL-CSG']
        list_str = []
        with open(file_path, "r") as source_f:
            for line in source_f:
                list_str.append(line.replace("\n", ""))

        # получение списка отчетов IP7270*
        list_ip = []
        counter = 0
        start = 0
        finish = 0
        for i, string in enumerate(list_str):
            if string.find("BUSINESS SERVICE ID SUBTOTAL") > -1:
                for counter, string2 in enumerate(list_str[i::-1]):
                    if string2.find("IP7270") == -1:
                        counter += 1
                    else:
                        break
                start = i - counter
                finish = i
                if list_str[start: start + finish] not in list_ip:
                    list_ip.append(list_str[start: finish])
                else:
                    break

        # получение отчетов, где есть типы операций FEE COLL-M2M, FEE COLL-MBG, FEE COLL-RET, FEE COLL-CSG
        tr_f_reports = []
        for lst in list_ip:
            for str_ip in lst:
                # print(str_ip)
                if str_ip.find(list_tr_f[0]) > -1 or str_ip.find(list_tr_f[1]) > -1 or str_ip.find(
                        list_tr_f[2]) > -1 or str_ip.find(list_tr_f[3]) > -1:
                    # print(lst)
                    tr_f_reports.append(lst)
                    break

        # выцепление данных из отчетов (до типов операций) для последующего занесения в excel файл
        dic_list = []
        for rep in tr_f_reports:
            dic = dict()
            for str_rep in rep:
                if str_rep.find("IP7270") > -1:
                    nspk_split = str_rep.split()
                    syst = " ".join(nspk_split[1:3])
                    dic.update({"system": syst})
                elif str_rep.find("CYCLE") > -1:
                    cycle_split = str_rep.split()
                    cycle = " ".join(cycle_split[3:8])
                    dic.update({"cycle": cycle})
                elif str_rep.find("BUSINESS SERVICE LEVEL") > -1:
                    date_split = str_rep.split()
                    str_date = date_split[4]
                    spl_date = str_date.split("-")
                    dd = spl_date[2]
                    mm = spl_date[1]
                    yyyy = spl_date[0]
                    date = f"{dd}.{mm}.{yyyy}"
                    dic.update({"date": date})
                elif str_rep.find("FILE ID") > -1:
                    file_id_split = str_rep.split()
                    file_id = " ".join(file_id_split[2:3])
                    dic.update({"file_id": file_id})
                elif str_rep.find("MEMBER ID") > -1:
                    member_id_split = str_rep.split()
                    member_id = " ".join(member_id_split[2:3])
                    c = 0
                    for symb in member_id:
                        if symb == '0':
                            c += 1
                        else:
                            break
                    dic.update({"member_id": member_id[c:]})
            dic_list.append(dic)

        # преобразование данных из каждого отчета (до операций) в отдельные списки для каждого отчета
        dic_of_dic_lists = {}
        for i, dictionary in enumerate(dic_list):
            system_list = []
            system_list.append(dictionary)
            dic_of_dic_lists.update({i: system_list})

        # получение списка операций
        index = 0
        ind_x = 0
        init = 0
        fin = 0
        tr_f_excel = []
        for rep in tr_f_reports:
            for index, st_ in enumerate(rep):
                if st_.find(list_tr_f[0]) > -1 or st_.find(list_tr_f[1]) > -1 or st_.find(
                        list_tr_f[2]) > -1 or st_.find(list_tr_f[3]) > -1:
                    init = index
                    break
            for ind_x, st_ in enumerate(rep[init:]):
                if st_.find(" TOTAL") > -1:
                    fin = ind_x + 1
                    break
            tr_f_excel.append(rep[init: init + fin])
            index = 0
            ind_x = 0

        # убрать -----
        i = 0
        for tr_f in tr_f_excel:
            for i, st in enumerate(tr_f):
                if st.find("---") == -1:
                    i += 1
                else:
                    break
            tr_f.pop(i)

        # выцепление данных по операциям для последующего занесения в excel-файл
        dic_list2 = {}
        dic_keys = ["trans_func", "proc_code", "counts", "recon_amount", "curr1", "pos1", "trans_fee", "curr2", "pos2"]
        for iter, tr_f in enumerate(tr_f_excel):
            some_lst = []
            res_lst = []
            for i, st in enumerate(tr_f):
                if not st[0:12].isspace():
                    trans_func = st[0:12]
                else:
                    st = tr_f[i - 1]
                    trans_func = st[0:12]
                    st = tr_f[i]
                some_lst.append(trans_func)
                proc_code = st[12:28].rstrip()
                some_lst.append(proc_code)
                counts = st[33:41].strip()
                some_lst.append(counts)
                recon_amount = st[41:65].strip()
                some_lst.append(recon_amount)
                curr1 = st[73:76].strip()
                some_lst.append(curr1)
                pos1 = st[65:68].strip()
                some_lst.append(pos1)
                trans_fee = st[76:97].strip()
                some_lst.append(trans_fee)
                curr2 = st[100:107]
                if curr2.strip() == "-":
                    curr2 = ""
                    some_lst.append(curr2)
                else:
                    some_lst.append(curr2[-3:])
                pos2 = st[97:100].strip()
                some_lst.append(pos2)
                res_dict = {}
                res_dict = dict(zip(dic_keys * (i + 1), some_lst))
                res_lst.append(res_dict)
            dic_list2.update({iter: res_lst})

        # итоговый словарь с отдельными списками для каждого отчета
        for k, v in dic_list2.items():
            tot_list = []
            for k2, v2 in dic_of_dic_lists.items():
                if k == k2:
                    for el in v:
                        dict_plus = {}
                        for el2 in v2:
                            dict_plus = el2 | el
                        tot_list.append(dict_plus)
            dic_list2.update({k: tot_list})

        # запись в файл excel
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # названия заголовков
        sheet["A1"] = "SYSTEM"
        sheet["B1"] = "CYCLE"
        sheet["C1"] = "DATE"
        sheet["D1"] = "FILE ID"
        sheet["E1"] = "MEMBER ID"
        sheet["F1"] = "TRANS. FUNC."
        sheet["G1"] = "PROC.CODE"
        sheet["H1"] = "COUNTS"
        sheet["I1"] = "RECON AMOUNT"
        sheet["J1"] = "CURR"
        sheet["K1"] = "POS"
        sheet["L1"] = "TRANS FEE"
        sheet["M1"] = "CURR"
        sheet["N1"] = "POS"

        # ширина столбцов
        sheet.column_dimensions["A"].width = 24
        sheet.column_dimensions["B"].width = 38
        sheet.column_dimensions["C"].width = 12
        sheet.column_dimensions["D"].width = 29
        sheet.column_dimensions["E"].width = 11
        sheet.column_dimensions["F"].width = 16
        sheet.column_dimensions["G"].width = 19
        sheet.column_dimensions["H"].width = 9
        sheet.column_dimensions["I"].width = 17
        sheet.column_dimensions["J"].width = 6
        sheet.column_dimensions["K"].width = 5
        sheet.column_dimensions["L"].width = 11
        sheet.column_dimensions["M"].width = 6
        sheet.column_dimensions["N"].width = 5

        # формат заголовков
        for cell in sheet["A:N"]:
            cell[0].font = Font(bold=True)
            cell[0].alignment = Alignment(horizontal="center")

        # заполнение ячеек из словаря (со второй строки) и настройка форматов
        row = 2
        for k, v in dic_list2.items():
            for dct in v:
                sheet[row][0].value = dct["system"]
                sheet[row][0].alignment = Alignment(horizontal="left")
                sheet[row][1].value = dct["cycle"]
                sheet[row][1].alignment = Alignment(horizontal="left")
                sheet[row][2].value = dct["date"]
                sheet[row][2].alignment = Alignment(horizontal="right")
                sheet[row][3].value = dct["file_id"]
                sheet[row][3].alignment = Alignment(horizontal="left")
                sheet[row][4].value = dct["member_id"]
                sheet[row][4].alignment = Alignment(horizontal="right")
                sheet[row][5].value = dct["trans_func"]
                sheet[row][5].alignment = Alignment(horizontal="left")
                sheet[row][6].value = dct["proc_code"]
                sheet[row][6].alignment = Alignment(horizontal="left")
                sheet[row][7].value = dct["counts"]
                sheet[row][7].alignment = Alignment(horizontal="right")
                sheet[row][8].value = dct["recon_amount"]
                sheet[row][8].alignment = Alignment(horizontal="right")
                sheet[row][9].value = dct["curr1"]
                sheet[row][9].alignment = Alignment(horizontal="left")
                sheet[row][10].value = dct["pos1"]
                sheet[row][10].alignment = Alignment(horizontal="left")
                sheet[row][11].value = dct["trans_fee"]
                sheet[row][11].alignment = Alignment(horizontal="right")
                sheet[row][12].value = dct["curr2"]
                sheet[row][12].alignment = Alignment(horizontal="left")
                sheet[row][13].value = dct["pos2"]
                sheet[row][13].alignment = Alignment(horizontal="left")
                row += 1
            row += 1

        workbook.save("pretense.xlsx")
        workbook.close()
        messagebox.showinfo("Готово!", "Программа выполнена")

open_button = Button(text="Выбрать файл", command=pretense, height=2, font="Arial 14")
open_button.config(background="#7FFFD4")
open_button.grid(row=2, padx=125,pady=65)
window.mainloop()