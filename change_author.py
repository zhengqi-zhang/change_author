import change_author_tools

# 文档作者处理系统
while True:

    # 显示功能菜单
    change_author_tools.show_menu()
    action_str = input("请选择希望执行的操作：")
    print(f"您选择的操作是【{action_str}】")

    # 1、2 针对文档的操作
    if action_str in ["1", "2"]:

        # 处理word文档作者信息
        if action_str == "1":
            change_author_tools.word_author()

        # 处理excel文档作者信息
        elif action_str == "2":
            change_author_tools.excel_author()

    # 0、退出系统
    elif action_str == "0":
        print("欢迎再次使用【文档作者信息处理系统】！")
        break

    # 其他内容输入错误，需要提示用户
    else:
        print("您选择有误，请重新选择")