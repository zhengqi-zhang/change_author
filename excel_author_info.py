import os
from openpyxl import load_workbook


# 判断用户输入的路径是否正确
def judgement_path(dir_path):
    """
	判断用户输入的路径是否存在
	:param dir_path:
	:return: dir_path
	:return: name
	"""
    if os.path.isdir(dir_path) or dir_path == '':
        if dir_path == '':
            dir_path = './'
        dir_path = dir_path.replace('\\', '/')
        return dir_path


def handle_file(dir_path, filepath_list=None, core_properties_list=None, wb_list=None):
    filepath_list = []
    core_properties_list = []
    wb_list = []
    for filename in os.listdir(dir_path):
        if filename.endswith('.xlsx') and not filename.startswith('~$'):
            filepath = os.path.join(dir_path, filename).replace('\\', '/')
            wb = load_workbook(filepath)
            core_properties = wb.properties
            filepath_list.append(filepath)
            core_properties_list.append(core_properties)
            wb_list.append(wb)
    return wb_list, filepath_list, core_properties_list


# 显示菜单
def excel_menu(dir_path):
    print("=" * 50)
    print("")
    print("1. 显示文档作者信息")
    print("2. 修改文档作者信息")
    print("")
    print("0. 退出")
    action_str = input("请选择希望执行的操作：")

    if action_str in ["1", "2"]:

        # 1、显示文档作者信息
        if action_str == "1":
            wb_list, filepath_list, core_properties_list = handle_file(dir_path)
            info_show(filepath_list, core_properties_list)

        # 2、修改文档作者信息
        elif action_str == "2":
            wb_list, filepath_list, core_properties_list = handle_file(dir_path)
            change_info(wb_list, filepath_list, core_properties_list)


    # 0、退出 .xlsx 文档作者信息处理
    elif action_str == "0":
        print("欢迎再次使用！")
        return "0"
    else:
        print("您选择有误，请重新选择！")


# 显示文档作者信息
def info_show(filepath_list, core_properties_list):
    print("显示文档作者信息")
    if core_properties_list:
        for filepath, core in zip(filepath_list, core_properties_list):
            print(f"{filepath} 文档的作者是 {core.creator}")
            print(f"{filepath} 文档最后的修改者是 {core.last_modified_by}")
    else:
        print(f'指定目录中没有 .xlsx 文档！')


# 修改文档作者信息
def change_info(wb_list, filepath_list, core_properties_list):
    print("修改文档作者信息")
    if filepath_list:
        name_c = input("请输入你想要的作者姓名(直接回车不修改)：")
        name_l = input("请输入你想要的最后修改者姓名(直接回车不修改)：")
        i = j = 0
        for wb, filepath, core in zip(wb_list, filepath_list, core_properties_list):
            if name_c != "":
                core.creator = name_c
                i += 1
            if name_l != "":
                core.last_modified_by = name_l
                j += 1
            if i > 0 or j > 0:
                wb.save(filepath)
        if i > 0:
            print(f'已修改指定目录中 {i} 个 .xlsx 文档的作者信息！')
        else:
            print(f'没有修改指定目录中 .xlsx 文档的作者信息！')
        if j > 0:
            print(f'已修改指定目录中 {j} 个 .xlsx 文档的最后修改者信息！')
        else:
            print(f'没有修改指定目录中 .xlsx 文档的最后修改者信息！')
    else:
        print(f'指定目录中没有 .xlsx 文档！')
