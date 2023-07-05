import os
from docx import Document


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


# 显示 .docx 文档处理菜单
def word_menu(dir_path):
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
			document_list, filepath_list, core_properties_list = handle_file(dir_path)
			info_show(filepath_list, core_properties_list)

		# 2、修改文档作者信息
		elif action_str == "2":
			document_list, filepath_list, core_properties_list = handle_file(dir_path)
			change_info(document_list, filepath_list, core_properties_list)

		# 0、退出 .docx 文档作者信息处理
	elif action_str == "0":
		print("欢迎再次使用！")
		return "0"
	else:
		print("您选择有误，请重新选择！")


def handle_file(dir_path, document_list=None, filepath_list=None, core_properties_list=None):
	filepath_list = []
	core_properties_list = []
	document_list = []
	for filename in os.listdir(dir_path):
		if filename.endswith('.docx') and not filename.startswith('~$'):
			filepath = os.path.join(dir_path, filename).replace('\\', '/')
			document = Document(filepath)
			core_properties = document.core_properties
			filepath_list.append(filepath)
			core_properties_list.append(core_properties)
			document_list.append(document)
	return document_list, filepath_list, core_properties_list


# 显示文档作者信息
def info_show(filepath_list, core_properties_list):
	print("显示文档作者信息")
	if core_properties_list:
		for filepath, core in zip(filepath_list, core_properties_list):
			print(f"{filepath} 文档的作者是 {core.author}")
			print(f"{filepath} 文档最后的修改者是 {core.last_modified_by}")
	else:
		print(f'指定目录中没有 .docx 文档！')


# 修改文档作者信息
def change_info(document_list, filepath_list, core_properties_list):
	print("修改文档作者信息")
	if filepath_list:
		name_a = input("请输入你想要的作者姓名(直接回车不修改)：")
		name_l = input("请输入你想要的最后修改者姓名(直接回车不修改)：")
		i = j = 0
		for document, filepath, core in zip(document_list, filepath_list, core_properties_list):
			if name_a != "":
				core.author = name_a
				i += 1
			if name_l != "":
				core.last_modified_by = name_l
				j += 1
			if i > 0 or j > 0:
				document.save(filepath)

		if i > 0:
			print(f'已修改指定目录中 {i} 个 .docx 文档的作者信息！')
		else:
			print(f'没有处理指定目录中 .docx 文档的作者信息！')
		if j > 0:
			print(f'已修改指定目录中 {j} 个 .docx 文档的最后修改者信息！')
		else:
			print(f'没有修改指定目录中 .docx 文档的最后修改者信息！')
	else:
		print(f'指定目录中没有 .docx 文档！')
