import pandas as pd

# -*- coding: utf-8 -*-

def transfer_data(excel_file):
    try:
        # 读取 Excel 文件中的所有表
        sheets = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
    except Exception as e:
        print(f"读取 Excel 文件时出错: {e}")
        return

    # 获取 '卡片管理' 工作表
    card_management_df = sheets.get('卡片管理')

    if card_management_df is None:
        print("没有找到 '卡片管理' 工作表。")
        return

    # 获取目标列（包括第一列）
    target_columns = list(card_management_df.columns)  # 包含第一列
    print(f"目标列: {target_columns}")

    # 遍历 '卡片管理' 数据框的每一行
    for index, row in card_management_df.iterrows():
        sheet_name = row[0]  # 第一列作为工作表名称
        data_to_send = row.tolist()  # 所有数据转换为列表

        # 判断“巡检类型”列是否包含“换”
        inspection_type = row.get('巡检类型')
        if pd.isna(inspection_type) or '换' not in str(inspection_type):
            print(f"行 {index + 1} 的 '巡检类型' 不包含 '换'，跳过此行。")
            continue

        # 打印当前处理的行信息
        print(f"处理行 {index + 1}: 目标表名='{sheet_name}' (类型: {type(sheet_name)}), 数据={data_to_send}")

        # 确保 sheet_name 是字符串且不为空
        if pd.isna(sheet_name):
            print(f"行 {index + 1} 的 sheet_name 是 NaN，跳过此行。")
            continue
        if not isinstance(sheet_name, str):
            sheet_name = str(sheet_name)
            print(f"目标表名转换为字符串: '{sheet_name}'")

        sheet_name = sheet_name.strip()  # 去除前后空格
        if not sheet_name:
            print(f"行 {index + 1} 的 sheet_name 为空字符串，跳过此行。")
            continue

        # 检查 sheet_name 是否包含 Excel 不支持的字符
        invalid_chars = [":", "\\", "/", "*", "?", "[", "]"]
        if any(char in sheet_name for char in invalid_chars):
            print(f"行 {index + 1} 的 sheet_name 包含无效字符 '{sheet_name}'，跳过此行。")
            continue

        # 检查目标表是否存在
        if sheet_name in sheets:
            # 获取目标数据框
            target_df = sheets[sheet_name]
            print(f"目标表 '{sheet_name}' 已存在，当前行数: {len(target_df)}")

            # 如果目标数据框没有列定义，初始化列名
            if target_df.empty:
                target_df = pd.DataFrame(columns=target_columns)
                print(f"目标表 '{sheet_name}' 为空，初始化列名。")
            else:
                # 检查目标表的列是否匹配
                if list(target_df.columns) != target_columns:
                    print(f"目标表 '{sheet_name}' 的列与源表不匹配。跳过此行: {data_to_send}")
                    continue
        else:
            # 如果目标表不存在，创建新表，列名与'卡片管理'表相同
            target_df = pd.DataFrame(columns=target_columns)
            print(f"目标表 '{sheet_name}' 不存在，创建新表。")

        # 检查插入的数据和目标数据框的列数是否匹配
        if len(data_to_send) != len(target_columns):
            print(f"数据和目标表列数不匹配，跳过此行: {data_to_send}")
            continue

        # 将数据添加到目标数据框
        try:
            target_df.loc[len(target_df)] = data_to_send
            print(f"成功插入数据到表 '{sheet_name}'。")
        except Exception as e:
            print(f"插入数据到表 '{sheet_name}' 时出错: {e}")
            continue

        # 更新目标表
        sheets[sheet_name] = target_df

    # 在保存之前，确保所有 sheet_names 都是有效字符串
    for sheet_name in list(sheets.keys()):
        if not isinstance(sheet_name, str):
            new_sheet_name = str(sheet_name)
            sheets[new_sheet_name] = sheets.pop(sheet_name)
            print(f"将 sheet_name '{sheet_name}' 转换为字符串 '{new_sheet_name}'。")

    # 打印所有 sheet_names 以调试
    print("所有 sheet_names:")
    for name in sheets.keys():
        print(f"'{name}' (类型: {type(name)})")

    # 保存更新后的工作簿
    try:
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                # 确保 sheet_name 是字符串
                if not isinstance(sheet_name, str):
                    sheet_name = str(sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print("数据已成功转移到相应的工作表。")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")

# 调用函数，传入 Excel 文件的路径
transfer_data('/python/巡检记录2024-10-18.xlsx')
