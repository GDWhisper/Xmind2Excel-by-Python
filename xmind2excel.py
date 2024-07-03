#Author:GDWhisper
#Date:2024-04-23


import openpyxl
import os
from xmindparser import xmind_to_dict



# 1.使用xmind_to_dict将xmind数据转为字典格式
def xmind_dict(filename):
    dict_data = xmind_to_dict(filename)
    # 输出topics下的子节点
    topics_data = dict_data[0]['topic']['topics']
    return topics_data


# 2.读取字典获取用例数据
def case_data(testdata):
    case_temp = []
    for big_title in testdata:  #获取大标题（大模块）
        for sub_title in big_title['topics']:   #获取子标题（子模块）
            for case_title in sub_title['topics']:  #获取用例标题
                for precondition in case_title['topics']:   #获取前置条件
                    level_temp=[precon['title'] for precon in precondition['topics'] if precon['title'] != 'step']   #获取用例优先级
                    level = ''.join(level_temp) #去括号
                    for step_title in precondition['topics']:
                        if step_title['title'] not in ('P0','P1','P2','P3','P4'):
                            step=[stepp['title'] for stepp in step_title['topics']]  #获取用例步骤
                            for result_t in step_title['topics']:
                                result_c = result_t.get('topics', [{}])[0].get('title', [{}])    #获取预期结果
                                if result_c != [{}]: #筛除空项，每条用例的预期结果只有一个
                                    result_case =result_c
                                    for result_title in result_t['topics']:
                                        remark = result_title.get('topics', [{}])[0].get('title', [{}]) #获取备注

                                        # 为迎合pingcode平台用例导入格式，大小标题使用"/"进行链接，如无需要可自行删改
                                        if remark != [{}]:
                                            case_temp.append((big_title['title']+ "/" +sub_title['title'], case_title['title'], precondition['title'],level,step,result_case,remark))
                                        else:
                                            remark = ""    #当备注节点缺失时，传入空值
                                            case_temp.append((big_title['title']+ "/" +sub_title['title'], case_title['title'],
                                                           precondition['title'], level, step, result_case, remark))
    return case_temp



# 3.根据xmind文件，获取excel文件地址及文件名，将excel文件保存到xmind文件所在目录
def excel_info(filename):
    # 表格基本信息
    excel_name = filename[filename.rindex('/') + 1:filename.index('.xmind')]
    excel_path = filename[:filename.rindex('/')]
    excel_file = os.path.join(excel_path, excel_name + '.xlsx')
    excel_file = get_available_filename(excel_file, '.xlsx')
    print('文件保存路径：',excel_file.replace('/', '\\'))
    return excel_file

# 检查文件是否存在，如果存在，则文件名后缀增加可以递增的数字
def get_available_filename(path, extension):
    filename, _ = os.path.splitext(path)
    if os.path.isfile(f"{filename}{extension}"):
        i = 1
        while os.path.isfile(f"{filename}({i}){extension}"):
            i += 1
        return f"{filename}({i}){extension}"
    else:
        return f"{filename}{extension}"



# 4.创建并将数据存储至excel中
def excel_data(excel_name, testdata):
    # 创建一个excel
    workbook = openpyxl.Workbook()
    # 获取当前活动的工作表
    worksheet = workbook.active
    # 设置表头
    headers = ['模块', '*标题', '前置条件', '重要程度', '步骤描述', '预期结果', '备注','维护人','用例类型','测试类型']
    worksheet.append(headers)
    # 输入case数据
    for row in case_data(xmind_dict(filename)):
        new_row = list(row)
        new_row.append(tester)  # 添加维护人列
        new_row.append(tctype)  # 添加用例类型列
        new_row.append(testtype)  # 添加测试类型列
        steps = new_row[4]  #提取测试步骤数据
        formatted_steps = []
        for i, step in enumerate(steps, start=1):   #开始处理测试步骤数据
            formatted_steps.append(f"{i}) {step}")  #给每个步骤添加序列号
        new_row[4] = "\n".join(formatted_steps) #给每个步骤分行并覆盖原步骤数据
        worksheet.append(new_row)

    ws = workbook.worksheets[0]
    ws.insert_rows(1)   # 为了迎合pingcode平台用例导入格式，在表头上面插入一行空白行，不需要可注释

    # 保存excel文件
    workbook.save(excel_name)


if __name__ == '__main__':
    filename = 'E:/地址/1.xmind'
    tester = '维护人'  # 维护人
    tctype = '功能测试'  # 用例类型
    testtype = '手动'  # 测试类型

    # while True:
    #     user_input = input("请输入xmind文件地址(包含文件格式)：")
    #     if user_input == '':
    #         break
    user_input = filename.replace('\\', '/')
    filename = user_input

    print("开始转换···")
    excel_data(excel_info(filename), xmind_dict(filename))
    print("执行结束！")

