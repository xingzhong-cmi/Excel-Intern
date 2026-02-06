"""
Excel merge and combination operations module
文件合并与组合函数
"""

import pandas as pd
import os
from typing import Tuple, List, Union


def excel_merge_files(file_paths: List[str], sheet_names: Union[str, List[str]], 
                     save_path: str, axis: int = 0) -> Tuple[bool, str]:
    """
    合并多个Excel文件
    
    Args:
        file_paths: Excel文件路径列表
        sheet_names: 工作表名或工作表名列表
        save_path: 保存路径
        axis: 合并方向 0=纵向（行拼接）/1=横向（列拼接）
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not file_paths:
            return False, "文件路径列表为空"
        
        # 检查文件是否存在
        for path in file_paths:
            if not os.path.exists(path):
                return False, f"文件不存在: {path}"
        
        # 读取所有文件
        dfs = []
        if isinstance(sheet_names, str):
            # 所有文件使用相同的工作表名
            for path in file_paths:
                df = pd.read_excel(path, sheet_name=sheet_names)
                dfs.append(df)
        elif isinstance(sheet_names, list):
            # 每个文件使用不同的工作表名
            if len(sheet_names) != len(file_paths):
                return False, f"工作表名数量({len(sheet_names)})与文件数量({len(file_paths)})不匹配"
            for path, sheet in zip(file_paths, sheet_names):
                df = pd.read_excel(path, sheet_name=sheet)
                dfs.append(df)
        else:
            return False, "sheet_names参数类型错误"
        
        # 合并数据
        if axis == 0:
            # 纵向合并（行拼接）
            merged_df = pd.concat(dfs, axis=0, ignore_index=True)
        elif axis == 1:
            # 横向合并（列拼接）
            merged_df = pd.concat(dfs, axis=1)
        else:
            return False, f"不支持的axis值: {axis}"
        
        # 保存结果
        merged_df.to_excel(save_path, index=False)
        
        return True, f"成功合并 {len(file_paths)} 个文件，结果保存至: {save_path}"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"合并文件失败: {str(e)}"


def excel_merge_sheets(file_path: str, sheet_names: List[str], save_path: str, 
                      axis: int = 0, new_sheet_name: str = "MergedSheet") -> Tuple[bool, str]:
    """
    合并同一Excel文件中的多个工作表
    
    Args:
        file_path: Excel文件路径
        sheet_names: 要合并的工作表名列表
        save_path: 保存路径
        axis: 合并方向 0=纵向/1=横向
        new_sheet_name: 新工作表名
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        if not sheet_names:
            return False, "工作表名列表为空"
        
        # 读取所有指定的工作表
        dfs = []
        for sheet in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet)
                dfs.append(df)
            except ValueError:
                return False, f"工作表 '{sheet}' 不存在"
        
        # 合并数据
        if axis == 0:
            merged_df = pd.concat(dfs, axis=0, ignore_index=True)
        elif axis == 1:
            merged_df = pd.concat(dfs, axis=1)
        else:
            return False, f"不支持的axis值: {axis}"
        
        # 保存结果
        merged_df.to_excel(save_path, sheet_name=new_sheet_name, index=False)
        
        return True, f"成功合并 {len(sheet_names)} 个工作表，结果保存至: {save_path}"
        
    except Exception as e:
        return False, f"合并工作表失败: {str(e)}"


def excel_join_files(left_file: str, right_file: str, left_sheet: str, right_sheet: str,
                    on_column: Union[str, List[str]], how: str = 'inner', 
                    save_path: str = None) -> Tuple[bool, Union[str, pd.DataFrame]]:
    """
    按列关联合并两个Excel文件（类似SQL JOIN）
    
    Args:
        left_file: 左表文件路径
        right_file: 右表文件路径
        left_sheet: 左表工作表名
        right_sheet: 右表工作表名
        on_column: 关联列名或列名列表
        how: 关联方式 'inner'/'left'/'right'/'outer'
        save_path: 保存路径，None表示返回DataFrame
        
    Returns:
        (成功/失败, 消息或DataFrame)
    """
    try:
        if not os.path.exists(left_file):
            return False, f"左表文件不存在: {left_file}"
        if not os.path.exists(right_file):
            return False, f"右表文件不存在: {right_file}"
        
        # 读取两个表
        left_df = pd.read_excel(left_file, sheet_name=left_sheet)
        right_df = pd.read_excel(right_file, sheet_name=right_sheet)
        
        # 检查关联列是否存在
        if isinstance(on_column, str):
            if on_column not in left_df.columns:
                return False, f"左表中不存在列: {on_column}"
            if on_column not in right_df.columns:
                return False, f"右表中不存在列: {on_column}"
        elif isinstance(on_column, list):
            for col in on_column:
                if col not in left_df.columns:
                    return False, f"左表中不存在列: {col}"
                if col not in right_df.columns:
                    return False, f"右表中不存在列: {col}"
        
        # 关联合并
        merged_df = pd.merge(left_df, right_df, on=on_column, how=how)
        
        if save_path:
            merged_df.to_excel(save_path, index=False)
            return True, f"成功关联合并，结果保存至: {save_path}"
        else:
            return True, merged_df
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"关联合并失败: {str(e)}"


def excel_append_data(base_file: str, append_file: str, base_sheet: str, 
                     append_sheet: str, save_path: str = None) -> Tuple[bool, str]:
    """
    将一个Excel文件的数据追加到另一个文件
    
    Args:
        base_file: 基础文件路径
        append_file: 要追加的文件路径
        base_sheet: 基础文件工作表名
        append_sheet: 追加文件工作表名
        save_path: 保存路径，None表示覆盖基础文件
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(base_file):
            return False, f"基础文件不存在: {base_file}"
        if not os.path.exists(append_file):
            return False, f"追加文件不存在: {append_file}"
        
        # 读取文件
        base_df = pd.read_excel(base_file, sheet_name=base_sheet)
        append_df = pd.read_excel(append_file, sheet_name=append_sheet)
        
        # 追加数据
        result_df = pd.concat([base_df, append_df], ignore_index=True)
        
        # 保存
        output_path = save_path if save_path else base_file
        result_df.to_excel(output_path, sheet_name=base_sheet, index=False)
        
        return True, f"成功追加 {len(append_df)} 行数据，总行数: {len(result_df)}"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"追加数据失败: {str(e)}"
