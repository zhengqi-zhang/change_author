import word_author_info
import excel_author_info


# 显示菜单
def show_menu():

	# 功能菜单
	print("*" * 50)
	print("欢迎使用【文档作者处理系统】V3.00")
	print("")
	print("1. 处理 .docx 格式文档")
	print("2. 处理 .xlxs 格式文档")
	print("")
	print("0. 退出系统")
	print("*" * 50)


def word_author():
	print('请输入 .docx 文档所在的绝对路径：')
	print('若直接回车，则默认指定当前目录。')
	dir_path = input()
	while True:
		dir_path = word_author_info.judgement_path(dir_path)
		action_str = word_author_info.word_menu(dir_path)
		if action_str == "0":
			break


def excel_author():
	print('请输入 .xlsx 文档所在的绝对路径：')
	print('若直接回车，则默认指定当前目录。')
	dir_path = input()
	while True:
		dir_path = excel_author_info.judgement_path(dir_path)
		action_str = excel_author_info.excel_menu(dir_path)
		if action_str == "0":
			break
