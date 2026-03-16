#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
作者信息:
姓名:  (Jitao)
邮箱: jitao@x'x'x'x'x'x'x'x.com
部门: -IC-工xxxxxxxx程组
公司: xxxxxx有限公司
创建日期: 20xx-03-16
版本: PROCESSER_v3.8.4.py

===============================================================================

Excel数据处理脚本 - 版本3.7（新增功能⑦和功能⑧）- 修改版：不输出空值统计
修改内容：
1. 功能①筛选条件改为第一列最靠下且不为空的值
2. 新增功能④：筛选第五列包含"xx"的记录，统计第二十五列各个值的数量（不统计空值）
3. 新增功能⑤：筛选第五列包含xx"的记录，检查第二十列并输出表格
4. 新增功能⑥：筛选第五列包含xx"的记录，统计第二十五列各个值的数量（不统计空值）
5. 新增功能⑦：筛选第五列包含xx"的记录，检查第二十列并输出表格
6. 新增功能⑧：筛选第五列包含xx"的记录，统计第二十五列各个值的数量（不统计空值）

主要修改：
1. 功能①：不再需要用户输入筛选值，自动选择第一列最靠下（最后一行开始向上）的非空值
2. 新增功能④：在功能③的基础上，统计第二十五列各个值的数量并输出分析报告（不统计空值）
3. 新增功能⑤：类似功能③但筛选xx"关键字
4. 新增功能⑥：类似功能④但筛选xx"关键字（不统计空值）
5. 新增功能⑦：类似功能⑤但筛选xx"关键字
6. 新增功能⑧：类似功能⑥但筛选xx"关键字（不统计空值）

功能：
1. 功能①：筛选第一列最靠下且不为空的值（从最后一行开始向上查找第一个非空值）
   检查筛选出来的记录的第六列有多少种不同的值，统计每个值的记录数量
   输出："本周新增问题a起，其中（F列中的每个值）有（他们对应的数量）起<分开写>"
2. 功能②：筛选出第一列包含"20xx"的记录，统计数量记为b
   检查筛选出来的记录的第五列，统计包含"0KM"或"field"的记录数量记为c，不包含的数量记为d
   输出："问题年总计问题b起，其中售后/0KMc起+产线d起"
3. 功能③：筛选出第五列包含"xx"的记录，统计数量记为nu
   检查筛选出来的记录的第二十列的值，统计ongoing记录数量为ongoin，非ongoing记录数量为co
   输出："一共nu 起，co pcs已完成（closed+completed），ongoin 起未关闭（ongoing）："
   输出六列表格：FA No.(第9列), Device(第6列), 分析进展(第13列), 反馈时间(第2列), 批次(第11列), 备注(第26列)
4. 功能④：筛选第五列包含"xx"的记录（与功能③相同）
   检查这些记录的第二十五列，统计各个值的数量（不统计空值）
   输出："xx问题根因分析"
   然后另起一行第二十五列各个值各有多少起，每一种单独一行（不显示空值统计）
5. 功能⑤：筛选出第五列包含xx"的记录，统计数量记为n
   检查筛选出来的记录的第二十列的值，统计ongoing记录数量为ongoi，非ongoing记录数量为c
   输出："一共n 起，c pcs已完成（closed+completed），ongoi 起未关闭（ongoing）："
   输出六列表格：FA No.(第9列), Device(第6列), 分析进展(第13列), 反馈时间(第2列), 批次(第11列), 备注(第26列)
6. 功能⑥：筛选第五列包含xx"的记录（与功能⑤相同）
   检查这些记录的第二十五列，统计各个值的数量（不统计空值）
   输出："xx问题根因分析"
   然后另起一行第二十五列各个值各有多少起，每一种单独一行（不显示空值统计）
7. 功能⑦：筛选出第五列包含xx"的记录，统计数量记为num
   检查筛选出来的记录的第二十列的值，统计ongoing记录数量为ongoing，非ongoing记录数量为com
   输出："一共num 起，com pcs已完成（closed+completed），ongoing 起未关闭（ongoing）："
   输出六列表格：FA No.(第9列), Device(第6列), 分析进展(第13列), 反馈时间(第2列), 批次(第11列), 备注(第26列)
8. 功能⑧：筛选第五列包含xx"的记录（与功能⑦相同）
   检查这些记录的第二十五列，统计各个值的数量（不统计空值）
   输出："xx问题根因分析"
   然后另起一行第二十五列各个值各有多少起，每一种单独一行（不显示空值统计）

使用说明：
1. 确保已安装pandas和openpyxl库：pip install pandas openpyxl
2. 运行脚本：python PROCESSER.py
"""

import pandas as pd
import os
import sys

def format_table_data(table_df):
    """
    格式化表格数据为可读的文本格式
    
    参数:
    table_df: 包含表格数据的DataFrame，列名已经是显示名称
    
    返回:
    str: 格式化后的表格字符串
    """
    if table_df.empty:
        return "无符合条件的记录"
    
    # 创建表头
    headers = list(table_df.columns)
    header_line = " | ".join(headers)
    separator = "-" * (len(header_line) + 10)
    
    # 创建表格行
    rows = []
    for _, row in table_df.iterrows():
        row_data = []
        for col in headers:
            value = row[col]
            # 处理空值
            if pd.isna(value):
                row_data.append('')
            elif str(value).strip() == '':
                row_data.append('')
            else:
                row_data.append(str(value).strip())
        rows.append(" | ".join(row_data))
    
    # 组合表格
    table = f"{header_line}\n{separator}\n" + "\n".join(rows)
    return table

def process_excel_file(file_path):
    """
    处理Excel文件的主要函数
    
    参数:
    file_path: Excel文件路径
    
    返回:
    dict: 包含功能①、功能②和功能③统计结果的字典
    """
    try:
        print(f"正在读取文件: {file_path}")
        
        # 尝试读取Excel文件，使用更灵活的表头检测
        try:
            # 先尝试有表头的方式读取
            df = pd.read_excel(file_path, dtype=str, header=0)  # 假设第一行是表头
            print("✓ 使用第一行作为表头")
        except Exception as e:
            print(f"尝试有表头读取失败: {str(e)}")
            # 尝试无表头方式读取
            try:
                df = pd.read_excel(file_path, dtype=str, header=None)  # 无表头
                print("✓ 使用无表头方式读取")
            except Exception as e2:
                print(f"无表头读取也失败: {str(e2)}")
                raise
        
        # 检查数据框是否有足够的列
        print(f"数据形状: {df.shape[0]} 行 × {df.shape[1]} 列")
        
        # 获取列名（根据读取方式）
        if df.columns[0] == 0 and isinstance(df.columns[0], int):  # 无表头，使用默认列名
            print("检测到无表头，使用默认列名")
            # 重置列名为Col_1, Col_2...
            df.columns = [f'Col_{i+1}' for i in range(len(df.columns))]
        
        # 显示前几列的列名，用于调试
        print("\n前10列的列名:")
        for i in range(min(10, len(df.columns))):
            print(f"  列{i+1} (索引{i}): '{df.columns[i]}'")
        
        # 获取各列的列名（基于0-based索引）
        col_names = {}
        max_col = min(30, df.shape[1])  # 检查前30列
        for i in range(max_col):
            col_names[i] = df.columns[i]
        
        print("-" * 60)
        
        # ==================== 功能①：筛选第一列最靠下且不为空的值 ====================
        print("\n" + "=" * 60)
        print("功能①：筛选第一列最靠下且不为空的值")
        print("=" * 60)
        
        first_col = col_names.get(0)
        sixth_col = col_names.get(5) if 5 in col_names else None
        
        if not sixth_col:
            print("错误：数据列数不足，无法执行功能①（需要至少6列）")
            func1_result = None
        else:
            # 获取第一列数据
            first_col_data = df[first_col].astype(str).str.strip()
            
            # 从最后一行开始向上查找第一个非空值
            bottom_value = None
            bottom_row_index = None
            
            # 从最后一行开始向上遍历
            for i in range(len(first_col_data) - 1, -1, -1):
                value = first_col_data.iloc[i]
                # 检查是否非空（不是空字符串且不是NaN）
                if value != '' and not pd.isna(value):
                    bottom_value = value
                    bottom_row_index = i
                    break
            
            if bottom_value is None:
                print("功能①：第一列所有值都为空，无法执行筛选")
                func1_result = None
            else:
                print(f"功能①：第一列最靠下的非空值（第{bottom_row_index + 1}行）: '{bottom_value}'")
                
                # 筛选第一列等于这个最靠下非空值的记录
                filtered_df_func1 = df[first_col_data == bottom_value]
                
                # 统计功能①筛选出的记录数量（a）
                a = len(filtered_df_func1)
                
                print(f"功能①：筛选出 {a} 条第一列值为 '{bottom_value}' 的记录")
                
                # 统计第六列的不同值及其数量
                sixth_col_values = filtered_df_func1[sixth_col].astype(str).str.strip()
                unique_values = sixth_col_values.unique()
                unique_values_count = len(unique_values)
                
                print(f"功能①：第六列共有 {unique_values_count} 种不同的值")
                
                # 统计第六列每个不同值的数量
                sixth_value_counts = sixth_col_values.value_counts().to_dict()
                
                print("功能①第六列各值的详细统计:")
                for value, count in sixth_value_counts.items():
                    print(f"  {value}: {count} 条记录")
                
                # 显示第一列最后几行的值，用于调试
                print(f"\n功能①：第一列最后5行的值（从下往上）:")
                last_rows_count = min(5, len(first_col_data))
                for i in range(len(first_col_data) - 1, len(first_col_data) - last_rows_count - 1, -1):
                    value = first_col_data.iloc[i]
                    status = "✓ 非空" if value != '' and not pd.isna(value) else "✗ 空"
                    print(f"  第{i+1}行: '{value}' ({status})")
                
                func1_result = {
                    'target_value': bottom_value,
                    'target_row_index': bottom_row_index,
                    'total_filtered': a,
                    'unique_values_count': unique_values_count,
                    'value_counts': sixth_value_counts,
                    'first_col_name': first_col,
                    'sixth_col_name': sixth_col
                }
        
        # ==================== 功能②：筛选第一列包含"20xx"的记录 ====================
        print("\n" + "=" * 60)
        print("功能②：筛选第一列包含'20xx'的记录")
        print("=" * 60)
        
        fifth_col = col_names.get(4) if 4 in col_names else None
        
        if not fifth_col:
            print("错误：数据列数不足，无法执行功能②（需要至少5列）")
            func2_result = None
        else:
            # 筛选第一列包含"20xx"的记录（不区分大xx）
            filtered_df_func2 = df[df[first_col].astype(str).str.contains('20xx', case=False, na=False)]
            
            # 统计功能②筛选出的记录数量（b）
            b = len(filtered_df_func2)
            
            if b == 0:
                print("功能②：没有找到第一列包含'20xx'的记录")
                func2_result = None
            else:
                print(f"功能②：筛选出 {b} 条第一列包含'20xx'的记录")
                
                # 获取第五列数据
                fifth_col_values = filtered_df_func2[fifth_col].astype(str).str.strip()
                
                # 统计第五列中包含"0KM"或"field"的记录数量（c）
                contains_0km_or_field = fifth_col_values.str.contains('0KM|field', case=False, na=False)
                c = contains_0km_or_field.sum()
                
                # 统计第五列中不包含"0KM"或"field"的记录数量（d）
                d = b - c
                
                print(f"功能②：第五列中包含'0KM'或'field'的记录数量(c): {c}")
                print(f"功能②：第五列中不包含'0KM'或'field'的记录数量(d): {d}")
                
                func2_result = {
                    'total_filtered_问题': b,
                    'contains_0km_or_field': c,
                    'not_contains_0km_or_field': d,
                    'fifth_col_name': fifth_col
                }
        
        # ==================== 功能③：筛选第五列包含"xx"的记录 ====================
        print("\n" + "=" * 60)
        print("功能③：筛选第五列包含'xx'的记录")
        print("=" * 60)
        
        # 检查是否有足够的列用于功能③
        required_cols = [4, 5, 19, 8, 12, 1, 10, 25]  # 第五列(4), 第六列(5), 第二十列(19), 第九列(8), 第二十一列(20), 第二列(1), 第十一列(10), 第二十六列(25)
        missing_cols = [i for i in required_cols if i not in col_names]
        
        if missing_cols:
            print(f"错误：数据列数不足，无法执行功能③")
            print(f"缺少的列索引: {missing_cols}")
            print(f"当前数据列数: {df.shape[1]}")
            print("请确保Excel文件至少有26列数据")
            func3_result = None
        else:
            # 显示功能③相关列的列名
            print("功能③相关列名:")
            print(f"  第五列 (索引4): '{col_names[4]}'")
            print(f"  第六列 (索引5): '{col_names[5]}'")
            print(f"  第九列 (索引8): '{col_names[8]}'")
            print(f"  第十一列 (索引10): '{col_names[10]}'")
            print(f"  第二十一列 (索引20): '{col_names[20]}'")
            print(f"  第二十列 (索引19): '{col_names[19]}'")
            print(f"  第二十六列 (索引25): '{col_names[25]}'")
            
            fifth_col = col_names[4]  # 第五列
            twentieth_col = col_names[19]  # 第二十列
            
            # 筛选第五列包含"xx"的记录（不区分大xx）
            filtered_df_func3 = df[df[fifth_col].astype(str).str.contains('xx', case=False, na=False)]
            
            # 统计功能③筛选出的记录数量（nu）
            nu = len(filtered_df_func3)
            
            if nu == 0:
                print("功能③：没有找到第五列包含'xx'的记录")
                # 显示第五列的前几个值，帮助用户确认
                print(f"第五列前10个不同的值: {df[fifth_col].astype(str).str.strip().unique()[:10]}")
                func3_result = None
            else:
                print(f"功能③：筛选出 {nu} 条第五列包含'xx'的记录")
                
                # 获取第二十列数据
                twentieth_col_values = filtered_df_func3[twentieth_col].astype(str).str.strip()
                
                # 统计第二十列值为"ongoing"的记录数量（ongoin）
                # 使用不区分大xx的匹配
                ongoing_mask = twentieth_col_values.str.contains('ongoing|on-going', case=False, na=False)
                ongoin = ongoing_mask.sum()
                
                # 统计第二十列值不为"ongoing"的记录数量（co）
                co = nu - ongoin
                
                print(f"功能③：第二十列为'ongoing'的记录数量: {ongoin}")
                print(f"功能③：第二十列不为'ongoing'的记录数量: {co}")
                
                # 准备表格数据（只包含第二十列为"ongoing"的记录）
                ongoing_records = filtered_df_func3[ongoing_mask]
                
                if len(ongoing_records) > 0:
                    # 提取表格所需列并重命名
                    table_data = ongoing_records[[
                        col_names[8],   # 第九列 -> FA No.
                        col_names[5],   # 第六列 -> Device
                        col_names[20],  # 第二十一列 -> 分析进展
                        col_names[1],   # 第二列 -> 反馈时间
                        col_names[10],  # 第十一列 -> 批次
                        col_names[25]   # 第二十六列 -> 备注
                    ]].copy()
                    
                    # 重命名列
                    table_data.columns = ['FA No.', 'Device', '分析进展', '反馈时间', '批次', '备注']
                    
                    # 格式化表格
                    formatted_table = format_table_data(table_data)
                    
                    func3_result = {
                        'nu': nu,
                        'ongoin': ongoin,
                        'co': co,
                        'fifth_col_name': fifth_col,
                        'twentieth_col_name': twentieth_col,
                        'table_data': table_data,
                        'formatted_table': formatted_table,
                        'ongoing_records_count': len(ongoing_records)
                    }
                else:
                    print("功能③：没有找到第二十列为'ongoing'的记录")
                    func3_result = None
        
        # ==================== 功能⑤：筛选第五列包含xx"的记录 ====================
        print("\n" + "=" * 60)
        print("功能⑤：筛选第五列包含xx'的记录")
        print("=" * 60)
        
        # 检查是否有足够的列用于功能⑤
        required_cols_func5 = [4, 5, 19, 8, 12, 1, 10, 25]  # 第五列(4), 第六列(5), 第二十列(19), 第九列(8), 第二十一列(20), 第二列(1), 第十一列(10), 第二十六列(25)
        missing_cols_func5 = [i for i in required_cols_func5 if i not in col_names]
        
        if missing_cols_func5:
            print(f"警告：数据列数不足，无法执行功能⑤")
            print(f"缺少的列索引: {missing_cols_func5}")
            print(f"当前数据列数: {df.shape[1]}")
            func5_result = None
        else:
            # 显示功能⑤相关列的列名
            print("功能⑤相关列名:")
            print(f"  第五列 (索引4): '{col_names[4]}'")
            print(f"  第六列 (索引5): '{col_names[5]}'")
            print(f"  第九列 (索引8): '{col_names[8]}'")
            print(f"  第十一列 (索引10): '{col_names[10]}'")
            print(f"  第二十一列 (索引20): '{col_names[20]}'")
            print(f"  第二十列 (索引19): '{col_names[19]}'")
            print(f"  第二十六列 (索引25): '{col_names[25]}'")
            
            fifth_col = col_names[4]  # 第五列
            twentieth_col = col_names[19]  # 第二十列
            
            # 筛选第五列包含xx"的记录（不区分大xx）
            filtered_df_func5 = df[df[fifth_col].astype(str).str.contains(xx', case=False, na=False)]
            
            # 统计功能⑤筛选出的记录数量（n）
            n = len(filtered_df_func5)
            
            if n == 0:
                print("功能⑤：没有找到第五列包含xx'的记录")
                func5_result = None
            else:
                print(f"功能⑤：筛选出 {n} 条第五列包含xx'的记录")
                
                # 获取第二十列数据
                twentieth_col_values = filtered_df_func5[twentieth_col].astype(str).str.strip()
                
                # 统计第二十列值为"ongoing"的记录数量（ongoi）
                # 使用不区分大xx的匹配
                ongoing_mask = twentieth_col_values.str.contains('ongoing|on-going', case=False, na=False)
                ongoi = ongoing_mask.sum()
                
                # 统计第二十列值不为"ongoing"的记录数量（c）
                c = n - ongoi
                
                print(f"功能⑤：第二十列为'ongoing'的记录数量: {ongoi}")
                print(f"功能⑤：第二十列不为'ongoing'的记录数量: {c}")
                
                # 准备表格数据（只包含第二十列为"ongoing"的记录）
                ongoing_records = filtered_df_func5[ongoing_mask]
                
                if len(ongoing_records) > 0:
                    # 提取表格所需列并重命名
                    table_data = ongoing_records[[
                        col_names[8],   # 第九列 -> FA No.
                        col_names[5],   # 第六列 -> Device
                        col_names[20],  # 第二十一列 -> 分析进展
                        col_names[1],   # 第二列 -> 反馈时间
                        col_names[10],  # 第十一列 -> 批次
                        col_names[25]   # 第二十六列 -> 备注
                    ]].copy()
                    
                    # 重命名列
                    table_data.columns = ['FA No.', 'Device', '分析进展', '反馈时间', '批次', '备注']
                    
                    # 格式化表格
                    formatted_table = format_table_data(table_data)
                    
                    func5_result = {
                        'n': n,
                        'ongoi': ongoi,
                        'c': c,
                        'fifth_col_name': fifth_col,
                        'twentieth_col_name': twentieth_col,
                        'table_data': table_data,
                        'formatted_table': formatted_table,
                        'ongoing_records_count': len(ongoing_records)
                    }
                else:
                    print("功能⑤：没有找到第二十列为'ongoing'的记录")
                    func5_result = None
        
        # ==================== 功能④：筛选第五列包含"xx"的记录，统计第二十五列各个值的数量 ====================
        print("\n" + "=" * 60)
        print("功能④：筛选第五列包含'xx'的记录，统计第二十五列各个值的数量（不统计空值）")
        print("=" * 60)
        
        # 检查是否有足够的列用于功能④（需要第二十五列，索引24）
        required_cols_func4 = [4, 24]  # 第五列(4), 第二十五列(24)
        missing_cols_func4 = [i for i in required_cols_func4 if i not in col_names]
        
        if missing_cols_func4:
            print(f"错误：数据列数不足，无法执行功能④")
            print(f"缺少的列索引: {missing_cols_func4}")
            print(f"当前数据列数: {df.shape[1]}")
            print("请确保Excel文件至少有25列数据")
            func4_result = None
        else:
            # 显示功能④相关列的列名
            print("功能④相关列名:")
            print(f"  第五列 (索引4): '{col_names[4]}'")
            print(f"  第二十五列 (索引24): '{col_names[24]}'")
            
            fifth_col = col_names[4]  # 第五列
            twenty_fifth_col = col_names[24]  # 第二十五列
            
            # 筛选第五列包含"xx"的记录（不区分大xx） - 与功能③相同
            filtered_df_func4 = df[df[fifth_col].astype(str).str.contains('xx', case=False, na=False)]
            
            # 统计功能④筛选出的记录数量（与功能③相同）
            nu_func4 = len(filtered_df_func4)
            
            if nu_func4 == 0:
                print("功能④：没有找到第五列包含'xx'的记录")
                func4_result = None
            else:
                print(f"功能④：筛选出 {nu_func4} 条第五列包含'xx'的记录")
                
                # 获取第二十五列数据
                twenty_fifth_col_values = filtered_df_func4[twenty_fifth_col].astype(str).str.strip()
                
                # 统计第二十五列各个值的数量
                # 处理空值或NaN值 - 只统计非空值
                # 先过滤掉空值和空字符串
                non_empty_values = twenty_fifth_col_values[~twenty_fifth_col_values.isna() & (twenty_fifth_col_values != '')]
                
                # 统计非空值的数量
                twenty_fifth_value_counts = non_empty_values.value_counts().to_dict()
                
                # 统计空值的数量（仅用于调试，不添加到结果中）
                empty_count = twenty_fifth_col_values.isna().sum() + (twenty_fifth_col_values == '').sum()
                if empty_count > 0:
                    print(f"功能④：第二十五列有 {empty_count} 条空值记录（不统计在输出结果中）")
                
                print(f"功能④：第二十五列共有 {len(twenty_fifth_value_counts)} 种不同的非空值")
                
                # 显示第二十五列各值的详细统计（非空值）
                print("功能④第二十五列各值的详细统计（仅非空值）:")
                total_count = 0
                for value, count in twenty_fifth_value_counts.items():
                    print(f"  '{value}': {count} 条记录")
                    total_count += count
                
                print(f"功能④：第二十五列非空值统计总数验证: {total_count} 条记录")
                if empty_count > 0:
                    print(f"功能④：第二十五列空值记录数量: {empty_count} 条（不输出）")
                
                func4_result = {
                    'nu': nu_func4,
                    'twenty_fifth_col_name': twenty_fifth_col,
                    'value_counts': twenty_fifth_value_counts,
                    'total_count': total_count,
                    'empty_count': empty_count
                }
        
        # ==================== 功能⑥：筛选第五列包含xx"的记录，统计第二十五列各个值的数量 ====================
        print("\n" + "=" * 60)
        print("功能⑥：筛选第五列包含xx'的记录，统计第二十五列各个值的数量（不统计空值）")
        print("=" * 60)
        
        # 检查是否有足够的列用于功能⑥（需要第二十五列，索引24）
        required_cols_func6 = [4, 24]  # 第五列(4), 第二十五列(24)
        missing_cols_func6 = [i for i in required_cols_func6 if i not in col_names]
        
        if missing_cols_func6:
            print(f"警告：数据列数不足，无法执行功能⑥")
            print(f"缺少的列索引: {missing_cols_func6}")
            print(f"当前数据列数: {df.shape[1]}")
            func6_result = None
        else:
            # 显示功能⑥相关列的列名
            print("功能⑥相关列名:")
            print(f"  第五列 (索引4): '{col_names[4]}'")
            print(f"  第二十五列 (索引24): '{col_names[24]}'")
            
            fifth_col = col_names[4]  # 第五列
            twenty_fifth_col = col_names[24]  # 第二十五列
            
            # 筛选第五列包含xx"的记录（不区分大xx）
            filtered_df_func6 = df[df[fifth_col].astype(str).str.contains(xx', case=False, na=False)]
            
            # 统计功能⑥筛选出的记录数量
            n_func6 = len(filtered_df_func6)
            
            if n_func6 == 0:
                print("功能⑥：没有找到第五列包含xx'的记录")
                func6_result = None
            else:
                print(f"功能⑥：筛选出 {n_func6} 条第五列包含xx'的记录")
                
                # 获取第二十五列数据
                twenty_fifth_col_values = filtered_df_func6[twenty_fifth_col].astype(str).str.strip()
                
                # 统计第二十五列各个值的数量
                # 处理空值或NaN值 - 只统计非空值
                # 先过滤掉空值和空字符串
                non_empty_values = twenty_fifth_col_values[~twenty_fifth_col_values.isna() & (twenty_fifth_col_values != '')]
                
                # 统计非空值的数量
                twenty_fifth_value_counts = non_empty_values.value_counts().to_dict()
                
                # 统计空值的数量（仅用于调试，不添加到结果中）
                empty_count = twenty_fifth_col_values.isna().sum() + (twenty_fifth_col_values == '').sum()
                if empty_count > 0:
                    print(f"功能⑥：第二十五列有 {empty_count} 条空值记录（不统计在输出结果中）")
                
                print(f"功能⑥：第二十五列共有 {len(twenty_fifth_value_counts)} 种不同的非空值")
                
                # 显示第二十五列各值的详细统计（非空值）
                print("功能⑥第二十五列各值的详细统计（仅非空值）:")
                total_count = 0
                for value, count in twenty_fifth_value_counts.items():
                    print(f"  '{value}': {count} 条记录")
                    total_count += count
                
                print(f"功能⑥：第二十五列非空值统计总数验证: {total_count} 条记录")
                if empty_count > 0:
                    print(f"功能⑥：第二十五列空值记录数量: {empty_count} 条（不输出）")
                
                func6_result = {
                    'n': n_func6,
                    'twenty_fifth_col_name': twenty_fifth_col,
                    'value_counts': twenty_fifth_value_counts,
                    'total_count': total_count,
                    'empty_count': empty_count
                }
        
        # ==================== 功能⑦：筛选第五列包含xx"的记录 ====================
        print("\n" + "=" * 60)
        print("功能⑦：筛选第五列包含xx'的记录")
        print("=" * 60)
        
        # 检查是否有足够的列用于功能⑦
        required_cols_func7 = [4, 5, 19, 8, 12, 1, 10, 25]  # 第五列(4), 第六列(5), 第二十列(19), 第九列(8), 第二十一列(20), 第二列(1), 第十一列(10), 第二十六列(25)
        missing_cols_func7 = [i for i in required_cols_func7 if i not in col_names]
        
        if missing_cols_func7:
            print(f"警告：数据列数不足，无法执行功能⑦")
            print(f"缺少的列索引: {missing_cols_func7}")
            print(f"当前数据列数: {df.shape[1]}")
            func7_result = None
        else:
            # 显示功能⑦相关列的列名
            print("功能⑦相关列名:")
            print(f"  第五列 (索引4): '{col_names[4]}'")
            print(f"  第六列 (索引5): '{col_names[5]}'")
            print(f"  第九列 (索引8): '{col_names[8]}'")
            print(f"  第十一列 (索引10): '{col_names[10]}'")
            print(f"  第二十一列 (索引20): '{col_names[20]}'")
            print(f"  第二十列 (索引19): '{col_names[19]}'")
            print(f"  第二十六列 (索引25): '{col_names[25]}'")
            
            fifth_col = col_names[4]  # 第五列
            twentieth_col = col_names[19]  # 第二十列
            
            # 筛选第五列包含xx"的记录（不区分大xx）
            filtered_df_func7 = df[df[fifth_col].astype(str).str.contains(xx', case=False, na=False)]
            
            # 统计功能⑦筛选出的记录数量（num）
            num = len(filtered_df_func7)
            
            if num == 0:
                print("功能⑦：没有找到第五列包含xx'的记录")
                func7_result = None
            else:
                print(f"功能⑦：筛选出 {num} 条第五列包含xx'的记录")
                
                # 获取第二十列数据
                twentieth_col_values = filtered_df_func7[twentieth_col].astype(str).str.strip()
                
                # 统计第二十列值为"ongoing"的记录数量（ongoing）
                # 使用不区分大xx的匹配
                ongoing_mask = twentieth_col_values.str.contains('ongoing|on-going', case=False, na=False)
                ongoing = ongoing_mask.sum()
                
                # 统计第二十列值不为"ongoing"的记录数量（com）
                com = num - ongoing
                
                print(f"功能⑦：第二十列为'ongoing'的记录数量: {ongoing}")
                print(f"功能⑦：第二十列不为'ongoing'的记录数量: {com}")
                
                # 准备表格数据（只包含第二十列为"ongoing"的记录）
                ongoing_records = filtered_df_func7[ongoing_mask]
                
                if len(ongoing_records) > 0:
                    # 提取表格所需列并重命名
                    table_data = ongoing_records[[
                        col_names[8],   # 第九列 -> FA No.
                        col_names[5],   # 第六列 -> Device
                        col_names[20],  # 第二十一列 -> 分析进展
                        col_names[1],   # 第二列 -> 反馈时间
                        col_names[10],  # 第十一列 -> 批次
                        col_names[25]   # 第二十六列 -> 备注
                    ]].copy()
                    
                    # 重命名列
                    table_data.columns = ['FA No.', 'Device', '分析进展', '反馈时间', '批次', '备注']
                    
                    # 格式化表格
                    formatted_table = format_table_data(table_data)
                    
                    func7_result = {
                        'num': num,
                        'ongoing': ongoing,
                        'com': com,
                        'fifth_col_name': fifth_col,
                        'twentieth_col_name': twentieth_col,
                        'table_data': table_data,
                        'formatted_table': formatted_table,
                        'ongoing_records_count': len(ongoing_records)
                    }
                else:
                    print("功能⑦：没有找到第二十列为'ongoing'的记录")
                    func7_result = None
        
        # ==================== 功能⑧：筛选第五列包含xx"的记录，统计第二十五列各个值的数量 ====================
        print("\n" + "=" * 60)
        print("功能⑧：筛选第五列包含xx'的记录，统计第二十五列各个值的数量（不统计空值）")
        print("=" * 60)
        
        # 检查是否有足够的列用于功能⑧（需要第二十五列，索引24）
        required_cols_func8 = [4, 24]  # 第五列(4), 第二十五列(24)
        missing_cols_func8 = [i for i in required_cols_func8 if i not in col_names]
        
        if missing_cols_func8:
            print(f"警告：数据列数不足，无法执行功能⑧")
            print(f"缺少的列索引: {missing_cols_func8}")
            print(f"当前数据列数: {df.shape[1]}")
            func8_result = None
        else:
            # 显示功能⑧相关列的列名
            print("功能⑧相关列名:")
            print(f"  第五列 (索引4): '{col_names[4]}'")
            print(f"  第二十五列 (索引24): '{col_names[24]}'")
            
            fifth_col = col_names[4]  # 第五列
            twenty_fifth_col = col_names[24]  # 第二十五列
            
            # 筛选第五列包含xx"的记录（不区分大xx）
            filtered_df_func8 = df[df[fifth_col].astype(str).str.contains(xx', case=False, na=False)]
            
            # 统计功能⑧筛选出的记录数量
            num_func8 = len(filtered_df_func8)
            
            if num_func8 == 0:
                print("功能⑧：没有找到第五列包含xx'的记录")
                func8_result = None
            else:
                print(f"功能⑧：筛选出 {num_func8} 条第五列包含xx'的记录")
                
                # 获取第二十五列数据
                twenty_fifth_col_values = filtered_df_func8[twenty_fifth_col].astype(str).str.strip()
                
                # 统计第二十五列各个值的数量
                # 处理空值或NaN值 - 只统计非空值
                # 先过滤掉空值和空字符串
                non_empty_values = twenty_fifth_col_values[~twenty_fifth_col_values.isna() & (twenty_fifth_col_values != '')]
                
                # 统计非空值的数量
                twenty_fifth_value_counts = non_empty_values.value_counts().to_dict()
                
                # 统计空值的数量（仅用于调试，不添加到结果中）
                empty_count = twenty_fifth_col_values.isna().sum() + (twenty_fifth_col_values == '').sum()
                if empty_count > 0:
                    print(f"功能⑧：第二十五列有 {empty_count} 条空值记录（不统计在输出结果中）")
                
                print(f"功能⑧：第二十五列共有 {len(twenty_fifth_value_counts)} 种不同的非空值")
                
                # 显示第二十五列各值的详细统计（非空值）
                print("功能⑧第二十五列各值的详细统计（仅非空值）:")
                total_count = 0
                for value, count in twenty_fifth_value_counts.items():
                    print(f"  '{value}': {count} 条记录")
                    total_count += count
                
                print(f"功能⑧：第二十五列非空值统计总数验证: {total_count} 条记录")
                if empty_count > 0:
                    print(f"功能⑧：第二十五列空值记录数量: {empty_count} 条（不输出）")
                
                func8_result = {
                    'num': num_func8,
                    'twenty_fifth_col_name': twenty_fifth_col,
                    'value_counts': twenty_fifth_value_counts,
                    'total_count': total_count,
                    'empty_count': empty_count
                }
        
        # 准备最终结果
        result = {
            'func1': func1_result,
            'func2': func2_result,
            'func3': func3_result,
            'func4': func4_result,
            'func5': func5_result,
            'func6': func6_result,
            'func7': func7_result,
            'func8': func8_result,
            'col_names': col_names,
            'total_rows': len(df),
            'total_cols': df.shape[1]
        }
        
        return result
        
    except FileNotFoundError:
        print(f"错误：找不到文件 '{file_path}'")
        return None
    except Exception as e:
        print(f"处理文件时发生错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def format_output(result):
    """
    格式化输出结果
    
    参数:
    result: 处理结果字典
    
    返回:
    str: 格式化后的输出字符串
    """
    if not result:
        return "没有找到符合条件的记录"
    
    output_lines = []
    
    # ==================== 功能①输出 ====================
    func1_result = result.get('func1')
    if func1_result:
        # 构建功能①输出字符串
        func1_output = f"本周新增问题{func1_result['total_filtered']}起，其中"
        
        # 添加每个值的统计
        value_items = []
        for value, count in func1_result['value_counts'].items():
            value_items.append(f"{value}有{count}起")
        
        func1_output += "；".join(value_items)
        output_lines.append(func1_output)
    else:
        output_lines.append("功能①：没有符合条件的记录")
    
    # ==================== 功能②输出 ====================
    func2_result = result.get('func2')
    if func2_result:
        # 构建功能②输出字符串
        func2_output = f"问题年总计问题{func2_result['total_filtered_问题']}起，其中售后/0KM{func2_result['contains_0km_or_field']}起+产线{func2_result['not_contains_0km_or_field']}起"
        output_lines.append(func2_output)
    else:
        output_lines.append("功能②：没有符合条件的记录")
    
    # ==================== 功能③输出 ====================
    func3_result = result.get('func3')
    if func3_result:
        # 构建功能③输出字符串
        func3_output = f"一共{func3_result['nu']} 起，{func3_result['co']} pcs已完成（closed+completed），{func3_result['ongoin']} 起未关闭（ongoing）："
        output_lines.append(func3_output)
        
        # 添加表格
        output_lines.append("\n" + "=" * 80)
        output_lines.append("xx相关未关闭记录表格：")
        output_lines.append("=" * 80)
        output_lines.append(func3_result['formatted_table'])
        
        # 添加表格统计信息
        output_lines.append(f"\n表格显示 {func3_result['ongoing_records_count']} 条记录")
    else:
        output_lines.append("功能③：没有符合条件的记录")
    
    # ==================== 功能④输出 ====================
    func4_result = result.get('func4')
    if func4_result:
        # 添加功能④标题
        output_lines.append("\nxx问题根因分析")
        
        # 添加第二十五列各个值的统计（只输出非空值）
        # 按照数量从高到低排序
        sorted_counts = sorted(func4_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
        
        for value, count in sorted_counts:
            # 只输出非空值，空值已不包含在value_counts中
            output_lines.append(f"{value}: {count}起")
        
        # 添加空值统计提示（可选）
        if func4_result.get('empty_count', 0) > 0:
            output_lines.append(f"（注：有{func4_result['empty_count']}条空值记录未计入统计）")
    else:
        output_lines.append("功能④：没有符合条件的记录")
    
    # ==================== 功能⑤输出 ====================
    func5_result = result.get('func5')
    if func5_result:
        # 构建功能⑤输出字符串
        func5_output = f"\n一共{func5_result['n']} 起，{func5_result['c']} pcs已完成（closed+completed），{func5_result['ongoi']} 起未关闭（ongoing）："
        output_lines.append(func5_output)
        
        # 添加表格
        output_lines.append("\n" + "=" * 80)
        output_lines.append("xx相关未关闭记录表格：")
        output_lines.append("=" * 80)
        output_lines.append(func5_result['formatted_table'])
        
        # 添加表格统计信息
        output_lines.append(f"\n表格显示 {func5_result['ongoing_records_count']} 条记录")
    else:
        output_lines.append("功能⑤：没有符合条件的记录")
    
    # ==================== 功能⑥输出 ====================
    func6_result = result.get('func6')
    if func6_result:
        # 添加功能⑥标题
        output_lines.append("\nxx问题根因分析")
        
        # 添加第二十五列各个值的统计（只输出非空值）
        # 按照数量从高到低排序
        sorted_counts = sorted(func6_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
        
        for value, count in sorted_counts:
            # 只输出非空值，空值已不包含在value_counts中
            output_lines.append(f"{value}: {count}起")
        
        # 添加空值统计提示（可选）
        if func6_result.get('empty_count', 0) > 0:
            output_lines.append(f"（注：有{func6_result['empty_count']}条空值记录未计入统计）")
    else:
        output_lines.append("功能⑥：没有符合条件的记录")
    
    # ==================== 功能⑦输出 ====================
    func7_result = result.get('func7')
    if func7_result:
        # 构建功能⑦输出字符串
        func7_output = f"\n一共{func7_result['num']} 起，{func7_result['com']} pcs已完成（closed+completed），{func7_result['ongoing']} 起未关闭（ongoing）："
        output_lines.append(func7_output)
        
        # 添加表格
        output_lines.append("\n" + "=" * 80)
        output_lines.append("x相关未关闭记录表格：")
        output_lines.append("=" * 80)
        output_lines.append(func7_result['formatted_table'])
        
        # 添加表格统计信息
        output_lines.append(f"\n表格显示 {func7_result['ongoing_records_count']} 条记录")
    else:
        output_lines.append("功能⑦：没有符合条件的记录")
    
    # ==================== 功能⑧输出 ====================
    func8_result = result.get('func8')
    if func8_result:
        # 添加功能⑧标题
        output_lines.append("\nx问题根因分析")
        
        # 添加第二十五列各个值的统计（只输出非空值）
        # 按照数量从高到低排序
        sorted_counts = sorted(func8_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
        
        for value, count in sorted_counts:
            # 只输出非空值，空值已不包含在value_counts中
            output_lines.append(f"{value}: {count}起")
        
        # 添加空值统计提示（可选）
        if func8_result.get('empty_count', 0) > 0:
            output_lines.append(f"（注：有{func8_result['empty_count']}条空值记录未计入统计）")
    else:
        output_lines.append("功能⑧：没有符合条件的记录")
    
    # 用换行符连接所有输出行
    return "\n".join(output_lines)

def main():
    """主函数"""
    print("=" * 80)
    print("Excel数据处理脚本 v3.7（新增功能⑦和功能⑧）- 修改版：不输出空值统计")
    print("修改内容：")
    print("  1. 功能①筛选条件改为第一列最靠下且不为空的值")
    print("  2. 新增功能④：筛选第五列包含'xx'的记录，统计第二十五列各个值的数量（不统计空值）")
    print("  3. 新增功能⑤：筛选第五列包含xx'的记录，检查第二十列并输出表格")
    print("  4. 新增功能⑥：筛选第五列包含xx'的记录，统计第二十五列各个值的数量（不统计空值）")
    print("  5. 新增功能⑦：筛选第五列包含xx'的记录，检查第二十列并输出表格")
    print("  6. 新增功能⑧：筛选第五列包含xx'的记录，统计第二十五列各个值的数量（不统计空值）")
    print("=" * 80)
    print("功能说明：")
    print("  1. 功能①：自动筛选第一列最靠下的非空值（从最后一行开始向上查找）")
    print("  2. 功能②：筛选第一列包含'20xx'的记录")
    print("  3. 功能③：筛选第五列包含'xx'的记录，输出未关闭记录表格")
    print("  4. 功能④：筛选第五列包含'xx'的记录，统计第二十五列各个值的数量（不统计空值）")
    print("  5. 功能⑤：筛选第五列包含xx'的记录，输出未关闭记录表格")
    print("  6. 功能⑥：筛选第五列包含xx'的记录，统计第二十五列各个值的数量（不统计空值）")
    print("  7. 功能⑦：筛选第五列包含xx'的记录，输出未关闭记录表格")
    print("  8. 功能⑧：筛选第五列包含xx'的记录，统计第二十五列各个值的数量（不统计空值）")
    print("=" * 80)
    print("主要修改：")
    print("  1. 功能①不再需要用户输入筛选值")
    print("  2. 自动从最后一行向上查找第一个非空值")
    print("  3. 新增功能④：xx问题根因分析（不统计空值）")
    print("  4. 新增功能⑤和功能⑥：xx问题分析（不统计空值）")
    print("  5. 新增功能⑦和功能⑧：x问题分析（不统计空值）")
    print("=" * 80)
    
    # 获取Excel文件路径
    while True:
        file_path = input("请输入Excel文件路径（或直接拖拽文件到终端）: ").strip().strip("'\"")
        
        if not file_path:
            print("文件路径不能为空")
            continue
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"文件 '{file_path}' 不存在，请重新输入")
            continue
        
        # 检查文件扩展名
        if not file_path.lower().endswith(('.xlsx', '.xls', '.csv')):
            print("警告：文件扩展名不是常见的Excel格式（.xlsx, .xls, .csv）")
            confirm = input("是否继续处理？(y/n): ").lower()
            if confirm != 'y':
                continue
        
        break
    
    print("\n" + "=" * 80)
    print("开始处理...")
    print("=" * 80)
    
    # 处理Excel文件（不再需要target_value参数）
    result = process_excel_file(file_path)
    
    print("\n" + "=" * 80)
    print("最终处理结果:")
    print("=" * 80)
    
    if result:
        # 显示基本信息
        print(f"文件信息:")
        print(f"  总行数: {result['total_rows']}")
        print(f"  总列数: {result['total_cols']}")
        
        # 显示各功能详细统计信息
        func1_result = result.get('func1')
        func2_result = result.get('func2')
        func3_result = result.get('func3')
        
        if func1_result:
            print(f"\n功能①详细统计:")
            print(f"  筛选条件：第一列最靠下的非空值 = '{func1_result['target_value']}'（第{func1_result['target_row_index'] + 1}行）")
            print(f"  筛选出的记录总数(a): {func1_result['total_filtered']}")
            print(f"  第六列不同值的数量: {func1_result['unique_values_count']}")
        
        if func2_result:
            print(f"\n功能②详细统计:")
            print(f"  筛选条件：第一列包含'20xx'")
            print(f"  筛选出的记录总数(b): {func2_result['total_filtered_问题']}")
            print(f"  第五列包含'0KM'或'field'的记录数量(c): {func2_result['contains_0km_or_field']}")
            print(f"  第五列不包含'0KM'或'field'的记录数量(d): {func2_result['not_contains_0km_or_field']}")
        
        if func3_result:
            print(f"\n功能③详细统计:")
            print(f"  筛选条件：第五列包含'xx'")
            print(f"  筛选出的记录总数(nu): {func3_result['nu']}")
            print(f"  第二十列为'ongoing'的记录数量(ongoin): {func3_result['ongoin']}")
            print(f"  第二十列不为'ongoing'的记录数量(co): {func3_result['co']}")
            print(f"  表格记录数: {len(func3_result['table_data'])}")
        
        # 显示功能④详细统计信息
        func4_result = result.get('func4')
        if func4_result:
            print(f"\n功能④详细统计:")
            print(f"  筛选条件：第五列包含'xx'")
            print(f"  筛选出的记录总数: {func4_result['nu']}")
            print(f"  第二十五列不同非空值的数量: {len(func4_result['value_counts'])}")
            if func4_result.get('empty_count', 0) > 0:
                print(f"  第二十五列空值记录数量: {func4_result['empty_count']}（不输出）")
            print(f"  第二十五列各值统计（仅非空值）:")
            # 按照数量从高到低排序
            sorted_counts = sorted(func4_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
            for value, count in sorted_counts:
                print(f"    '{value}': {count} 条记录")
        
        # 显示功能⑤详细统计信息
        func5_result = result.get('func5')
        if func5_result:
            print(f"\n功能⑤详细统计:")
            print(f"  筛选条件：第五列包含xx'")
            print(f"  筛选出的记录总数(n): {func5_result['n']}")
            print(f"  第二十列为'ongoing'的记录数量(ongoi): {func5_result['ongoi']}")
            print(f"  第二十列不为'ongoing'的记录数量(c): {func5_result['c']}")
            print(f"  表格记录数: {len(func5_result['table_data'])}")
        
        # 显示功能⑥详细统计信息
        func6_result = result.get('func6')
        if func6_result:
            print(f"\n功能⑥详细统计:")
            print(f"  筛选条件：第五列包含xx'")
            print(f"  筛选出的记录总数: {func6_result['n']}")
            print(f"  第二十五列不同非空值的数量: {len(func6_result['value_counts'])}")
            if func6_result.get('empty_count', 0) > 0:
                print(f"  第二十五列空值记录数量: {func6_result['empty_count']}（不输出）")
            print(f"  第二十五列各值统计（仅非空值）:")
            # 按照数量从高到低排序
            sorted_counts = sorted(func6_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
            for value, count in sorted_counts:
                print(f"    '{value}': {count} 条记录")
        
        # 显示功能⑦详细统计信息
        func7_result = result.get('func7')
        if func7_result:
            print(f"\n功能⑦详细统计:")
            print(f"  筛选条件：第五列包含xx'")
            print(f"  筛选出的记录总数(num): {func7_result['num']}")
            print(f"  第二十列为'ongoing'的记录数量(ongoing): {func7_result['ongoing']}")
            print(f"  第二十列不为'ongoing'的记录数量(com): {func7_result['com']}")
            print(f"  表格记录数: {len(func7_result['table_data'])}")
        
        # 显示功能⑧详细统计信息
        func8_result = result.get('func8')
        if func8_result:
            print(f"\n功能⑧详细统计:")
            print(f"  筛选条件：第五列包含xx'")
            print(f"  筛选出的记录总数: {func8_result['num']}")
            print(f"  第二十五列不同非空值的数量: {len(func8_result['value_counts'])}")
            if func8_result.get('empty_count', 0) > 0:
                print(f"  第二十五列空值记录数量: {func8_result['empty_count']}（不输出）")
            print(f"  第二十五列各值统计（仅非空值）:")
            # 按照数量从高到低排序
            sorted_counts = sorted(func8_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
            for value, count in sorted_counts:
                print(f"    '{value}': {count} 条记录")
        
        print("\n" + "=" * 80)
        print("格式化输出:")
        print("=" * 80)
        
        # 格式化输出
        formatted_output = format_output(result)
        print(formatted_output)
        
        # 可选：将结果保存到文件
        save_option = input("\n是否将结果保存到文件？(y/n): ").lower()
        if save_option == 'y':
            output_file = input("请输入输出文件名（默认: output_v3.7_no_empty.txt）: ").strip()
            if not output_file:
                output_file = "output_v3.7_no_empty.txt"
            
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write("Excel数据处理报告 v3.7（新增功能⑦和功能⑧）- 不统计空值版\n")
                    f.write("=" * 60 + "\n\n")
                    f.write(formatted_output + "\n")
                print(f"结果已保存到 '{output_file}'")
                
                # 同时保存详细统计信息
                detail_file = output_file.replace('.txt', '_details.txt') if '.txt' in output_file else output_file + '_details.txt'
                with open(detail_file, 'w', encoding='utf-8') as f:
                    f.write("Excel数据处理详细统计报告 v3.7（不统计空值版）\n")
                    f.write("=" * 60 + "\n\n")
                    f.write(f"处理文件: {file_path}\n")
                    f.write(f"处理时间: 问题年3月16日\n")
                    f.write(f"总行数: {result['total_rows']}\n")
                    f.write(f"总列数: {result['total_cols']}\n\n")
                    
                    if func1_result:
                        f.write("功能①统计:\n")
                        f.write(f"- 筛选条件: 第一列最靠下的非空值 = '{func1_result['target_value']}'（第{func1_result['target_row_index'] + 1}行）\n")
                        f.write(f"- 记录总数(a): {func1_result['total_filtered']}\n")
                        f.write(f"- 第六列不同值数量: {func1_result['unique_values_count']}\n")
                        f.write("- 第六列各值统计:\n")
                        for value, count in func1_result['value_counts'].items():
                            f.write(f"  * {value}: {count} 条记录\n")
                    
                    if func2_result:
                        f.write("\n功能②统计:\n")
                        f.write(f"- 筛选条件: 第一列包含'20xx'\n")
                        f.write(f"- 记录总数(b): {func2_result['total_filtered_问题']}\n")
                        f.write(f"- 包含'0KM'或'field'的记录(c): {func2_result['contains_0km_or_field']}\n")
                        f.write(f"- 不包含'0KM'或'field'的记录(d): {func2_result['not_contains_0km_or_field']}\n")
                    
                    if func3_result:
                        f.write("\n功能③统计:\n")
                        f.write(f"- 筛选条件: 第五列包含'xx'\n")
                        f.write(f"- 记录总数(nu): {func3_result['nu']}\n")
                        f.write(f"- 'ongoing'记录数量: {func3_result['ongoin']}\n")
                        f.write(f"- 非'ongoing'记录数量: {func3_result['co']}\n")
                        f.write(f"- 表格记录数: {len(func3_result['table_data'])}\n")
                    
                    if func4_result:
                        f.write("\n功能④统计:\n")
                        f.write(f"- 筛选条件: 第五列包含'xx'\n")
                        f.write(f"- 记录总数: {func4_result['nu']}\n")
                        f.write(f"- 第二十五列不同非空值数量: {len(func4_result['value_counts'])}\n")
                        if func4_result.get('empty_count', 0) > 0:
                            f.write(f"- 第二十五列空值记录数量: {func4_result['empty_count']}（不输出）\n")
                        f.write("- 第二十五列各值统计（仅非空值）:\n")
                        # 按照数量从高到低排序
                        sorted_counts = sorted(func4_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
                        for value, count in sorted_counts:
                            f.write(f"  * '{value}': {count} 条记录\n")
                    
                    if func5_result:
                        f.write("\n功能⑤统计:\n")
                        f.write(f"- 筛选条件: 第五列包含xx'\n")
                        f.write(f"- 记录总数(n): {func5_result['n']}\n")
                        f.write(f"- 'ongoing'记录数量: {func5_result['ongoi']}\n")
                        f.write(f"- 非'ongoing'记录数量: {func5_result['c']}\n")
                        f.write(f"- 表格记录数: {len(func5_result['table_data'])}\n")
                    
                    if func6_result:
                        f.write("\n功能⑥统计:\n")
                        f.write(f"- 筛选条件: 第五列包含xx'\n")
                        f.write(f"- 记录总数: {func6_result['n']}\n")
                        f.write(f"- 第二十五列不同非空值数量: {len(func6_result['value_counts'])}\n")
                        if func6_result.get('empty_count', 0) > 0:
                            f.write(f"- 第二十五列空值记录数量: {func6_result['empty_count']}（不输出）\n")
                        f.write("- 第二十五列各值统计（仅非空值）:\n")
                        # 按照数量从高到低排序
                        sorted_counts = sorted(func6_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
                        for value, count in sorted_counts:
                            f.write(f"  * '{value}': {count} 条记录\n")
                    
                    if func7_result:
                        f.write("\n功能⑦统计:\n")
                        f.write(f"- 筛选条件: 第五列包含xx'\n")
                        f.write(f"- 记录总数(num): {func7_result['num']}\n")
                        f.write(f"- 'ongoing'记录数量: {func7_result['ongoing']}\n")
                        f.write(f"- 非'ongoing'记录数量: {func7_result['com']}\n")
                        f.write(f"- 表格记录数: {len(func7_result['table_data'])}\n")
                    
                    if func8_result:
                        f.write("\n功能⑧统计:\n")
                        f.write(f"- 筛选条件: 第五列包含xx'\n")
                        f.write(f"- 记录总数: {func8_result['num']}\n")
                        f.write(f"- 第二十五列不同非空值数量: {len(func8_result['value_counts'])}\n")
                        if func8_result.get('empty_count', 0) > 0:
                            f.write(f"- 第二十五列空值记录数量: {func8_result['empty_count']}（不输出）\n")
                        f.write("- 第二十五列各值统计（仅非空值）:\n")
                        # 按照数量从高到低排序
                        sorted_counts = sorted(func8_result['value_counts'].items(), key=lambda x: x[1], reverse=True)
                        for value, count in sorted_counts:
                            f.write(f"  * '{value}': {count} 条记录\n")
                
                print(f"详细统计信息已保存到 '{detail_file}'")
                
            except Exception as e:
                print(f"保存文件时出错: {str(e)}")
    
    print("\n" + "=" * 80)
    print("处理完成！")
    print("=" * 80)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
        sys.exit(0)
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
        sys.exit(1)