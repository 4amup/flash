import xlwings as xw
import os
import sys
import time

# 1.创建输出文件夹的路径
cur_dir_name = time.strftime('%Y%m%d%H%M%S', time.localtime(
    time.time()))  # 格式化时间字符串：如20190822080449
desk_path = os.path.join(os.path.expanduser("~"), 'Desktop')  # 当前用户桌面路径
output_path = os.path.join(desk_path, cur_dir_name)  # 创建文件夹

# 2.创建文件夹
if os.path.exists(output_path):
    pass
else:
    os.mkdir(output_path)  # 不存在则在桌面创建一个文件夹

# 获取拖动文件的路径
try:
    drag_file_path = sys.argv[1]  # 拖动文件到exe可执行程序上时，第二个参数是拖上来的文件的绝对路径
except Exception as e:
    f = open(output_path + './log.txt', 'w')
    f.write(str(e)+'源文件传入错误')
finally:
    drag_file_path = r'C:/Users/lee/Desktop/66.XLSX'

# 2.打开文件并获取相关信息到数组中
app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False
wb = app.books.open(drag_file_path)
sht = wb.sheets[0]

# 3.取表头和整表数据
table_head_list = sht.range('A1').expand('right')
table_data = sht.range('A2').expand()

# 4.找供应商所在列SUPPLIER_COLUMN


def find_supplier_col(head_list):
    for item in head_list:
        if item.value == '供应商名称':
            return item.column


SUPPLIER_COLUMN = find_supplier_col(table_head_list)

# 5.过滤供应商为空的行


def supplier_empty(l):
    return l[SUPPLIER_COLUMN-1] and l[SUPPLIER_COLUMN-1].strip()


supplier_filter = list(filter(supplier_empty, table_data.value))

# 6.按供应商排序


def supplier_name_sort(l):
    return l[SUPPLIER_COLUMN-1]


supplier_sorted = sorted(supplier_filter, key=supplier_name_sort)

# 7.取供应商形成一个排序后列表
supplier_name_list = []
for item in supplier_sorted:
    supplier_name_list.append(item[SUPPLIER_COLUMN-1])

supplier_name_list = list(set(supplier_name_list))  # 数组去重
supplier_name_list.sort()

# 3.处理数组：筛选（去掉供应商列为空的行），按【供应商】排序，按供应商拆分文件
# 算法逻辑，取变化点的index，然后倒置，拆分长数组


def find_split_idx(suppliers):
    supplier_name_temp = ''
    i = 0
    arr = []
    for l in suppliers:
        supplier_name = l[SUPPLIER_COLUMN-1]
        if supplier_name_temp != supplier_name:
            supplier_name_temp = supplier_name
            arr.append(i)
        i += 1
    return arr


split_idx = find_split_idx(supplier_sorted)
split_idx.reverse()


def split_data(supplier_list):
    data = []
    arr = supplier_list
    for l in split_idx:
        group = arr[l:]
        arr = arr[:l]
        data.append(group)
    data.reverse()
    return data


all_data = split_data(supplier_sorted)

# 4.创建文件夹并保存拆分的文件
counter = 0
for data in all_data:
    try:
        data.insert(0, table_head_list.value)  # 插入表头
        temp_wb = xw.Book()
        temp_wb.sheets['sheet1'].range('A1').value = data
        temp_wb.save(os.path.join(
            output_path, supplier_name_list[counter]+'.XLSX'))
        temp_wb.close()
        counter += 1
    except Exception as e:
        f = open(output_path + './log.txt', 'w')
        f.write(str(e) + '-格式异常，请检查相关单元格')
        f.close()
        if temp_wb:
            temp_wb.close()

# 退出文件
if wb:
    wb.close()
if app:
    app.quit()
