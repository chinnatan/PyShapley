import pathlib
from datetime import datetime

import pandas as pd

import math

import statsmodels.api as sm

import itertools

import sys

# Global Variable
msg_pls_enter_data = 'Please enter data'
msg_pls_selected_excel_file_before_cal = 'Please select excel file before calculate'
msg_excel_file_only = 'Please select excel file (.xlsx) only'


def readFileExcel(filename):
    if filename is None:
        return None
    elif filename == '':
        return ''

    dataframe = pd.read_excel(filename)
    return dataframe


def writeFileExcel(shapley_dict):
    try:
        if shapley_dict is None or len(shapley_dict) < 0:
            raise Exception(msg_pls_enter_data)

        x_interest = []
        shapley_value = []
        key_of_shapley_dict = list(shapley_dict.keys())

        for key in key_of_shapley_dict:
            x_interest.append(key)
            shapley_value.append(shapley_dict.get(key))

        df = pd.DataFrame({'Variable "X" of Interest': x_interest, 'Shapley Value': shapley_value})

        current_date = datetime.now()
        current_date = current_date.strftime("%d%m%Y%H%M%S")
        current_path = pathlib.Path(__file__).parent.absolute()
        filename = file_path.split('\\')
        filename = filename[len(filename) - 1].split('.')
        filename_for_writer = "{path}\\{name}_calculate_{date}.xlsx".format(path=current_path, name=filename[0],
                                                                            date=current_date)

        df.to_excel(filename_for_writer, sheet_name='{name}_calculate'.format(name=filename[0]), index=False,
                    header=True)
    except Exception as ex:
        print(ex)


def combinations(func_x_of_head_list):
    if not len(func_x_of_head_list):
        return msg_pls_enter_data

    comb_list = []
    for i in range(1, len(func_x_of_head_list) + 1):
        for j in itertools.combinations(func_x_of_head_list, i):
            comb_list.append(j)
    return comb_list


def manageXofHeadList(x_var_list):
    if not len(x_var_list):
        return msg_pls_enter_data

    result_x_list = []
    for x_var in x_var_list:
        _x = []
        for j in x_var:
            _x.append(j)
        result_x_list.append(_x)
    return result_x_list


def calRSquare(x_var_list, main_data_var, main_y_var):
    if not len(x_var_list):
        return msg_pls_enter_data

    x_of_rsquare_dict = {}

    for i in x_var_list:
        key = ''.join(i)
        x_var = main_data_var[i]
        y_var = main_y_var
        sm_x_var = sm.add_constant(x_var)
        mlr_model = sm.OLS(y_var, sm_x_var)
        mlr_reg = mlr_model.fit()
        x_of_rsquare_dict[key] = mlr_reg.rsquared
    return x_of_rsquare_dict


# Function for calculate shapley value
def calculateShapley(filepath):
    try:
        if filepath is None or filepath == '':
            raise Exception(msg_pls_selected_excel_file_before_cal)
        elif filepath.find('xlsx') < 0:
            raise Exception(msg_excel_file_only)

        data = readFileExcel(filepath)

        y = data['Y']

        data.drop(['Y'], axis=1, inplace=True)  # ลบคอลัมน์ Y ออกจาก data

        head = list(data.head())
        print(f'[Log] All variable is {head}')

        # หาค่าความเป็นไปได้ของ X ว่ามีได้กี่รูปแบบ
        x_of_head_list = combinations(head)
        # print(f'[Log] Combination - {x_of_head_list}')

        # จัดระเรียบความเรียบร้อยของ X
        x_list = manageXofHeadList(x_of_head_list)
        # print(f'[Log] X List is {x_list}')

        # คำนวณค่า r-square
        x_of_rsquare_dict = calRSquare(x_list, data, y)
        # print(f'[Log] RSquare of X is {x_of_rsquare_dict}')

        # Main Process
        x_list_for_match = x_list
        head_list = head

        shapley_value_dict = {}

        for head in head_list:
            print(f'[Log] Interest in {head}')
            # กำหนดค่าให้ตัวแปรใหม่ กรณีเปลี่ยน X ที่สนใจ
            r = len(head_list)
            valid_a_match_list = []
            valid_b_match_list = []
            valid_c_match_list = []

            for count in range(1, r + 1):
                for x in x_list_for_match:
                    if len(x) == r and r != 1:
                        # ตรวจสอบค่า x ว่ามีใน list b หรือไม่ ถ้ายังไม่มีให้เพิ่มเข้า list a
                        if x not in valid_b_match_list:
                            valid_a_match_list.append(x)
                    elif len(x) == r and r == 1:  # กรณี r = 1 ให้เพิ่มเข้า list c
                        if x not in valid_a_match_list and x not in valid_b_match_list:
                            valid_c_match_list.append(x)
                for x in x_list_for_match:
                    if len(x) == r - 1 and head not in x:
                        # ตรวจสอบค่า x ว่ามีใน list a หรือไม่ ถ้ายังไม่มีให้เพิ่มเข้า list b
                        if x not in valid_a_match_list:
                            valid_b_match_list.append(x)
                r = r - 1
            # print(f'[Log] {head} list a is {valid_a_match_list}')
            # print(f'[Log] {head} list b is {valid_b_match_list}')
            # print(f'[Log] {head} list c is {valid_c_match_list}')

            # ทำการจับคู่ และคำนวณค่า k ไว้เบื้องต้น
            flag = 0
            match_a_into_b = []
            match_for_k = []

            for i in range(len(valid_a_match_list)):
                new_flag = len(valid_a_match_list[i])  # กำหนด flag เพื่อตรวจสอบว่าเป็น step เดียวกันไหม
                if flag == 0:
                    flag = new_flag
                if flag != new_flag:
                    match_a_into_b.append(match_for_k)
                    flag = new_flag
                    match_for_k = [[''.join(valid_a_match_list[i]), ''.join(valid_b_match_list[i])]]
                elif flag == new_flag:
                    group = [''.join(valid_a_match_list[i]), ''.join(valid_b_match_list[i])]
                    match_for_k.append(group)
                # ถ้า i เท่ากับตัวสุดท้ายของ list ให้เพิ่มข้อมูลเข้า list หลักและจบลูปพอดี
                if i == (len(valid_a_match_list) - 1):
                    match_a_into_b.append(match_for_k)
            # เพิ่มข้อมูลตัวท้ายสุดเข้า list หลัก
            match_a_into_b.append([[''.join(valid_c_match_list[0])]])
            # print(f'[Log] {head} Match a and b is {match_a_into_b}')

            # คำนวณค่า Shapley value ตาม X ที่สนใจ
            m = len(head_list)  # m = จำนวนตัวแปรอิสระทั้งหมดในสมการต้นแบบ
            shapley_value_total = 0

            for match in match_a_into_b:
                k = len(match)
                for data in match:
                    r_square_1 = x_of_rsquare_dict.get(data[0])
                    r_square_2 = x_of_rsquare_dict.get(data[len(data) - 1]) if (len(data) - 1) > 0 else 0
                    shapley_value_total += math.fabs((r_square_1 - r_square_2) / k)
            shapley_value_dict[head] = shapley_value_total / m
            print(f'[Log] Total Shapley Value is {shapley_value_dict[head]}')
        return shapley_value_dict
    except Exception as ex:
        print(ex)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    file_path = sys.argv[1:][0] if (len(sys.argv[1:])) > 0 else None

    shapley = calculateShapley(file_path)

    writeFileExcel(shapley)
