"""
Excel CRUD operations module
增删改查基础函数
"""

import pandas as pd
import os
from typing import Union, Tuple, List


def excel_add_row(file_path: str, sheet_name: str, row_data: dict, save_path: str = None) -> Tuple[bool, str]:
    """
    新增行到Excel文件
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        row_data: 要添加的行数据，字典格式 {列名: 值}
        save_path: 保存路径，默认为None时覆盖原文件
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 添加新行
        new_row = pd.DataFrame([row_data])
        df = pd.concat([df, new_row], ignore_index=True)
        
        # 保存文件
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功添加行到 {sheet_name}"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"添加行失败: {str(e)}"


def excel_add_column(file_path: str, sheet_name: str, column_name: str, 
                     column_data: list = None, default_value=None, save_path: str = None) -> Tuple[bool, str]:
    """
    新增列到Excel文件
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 新列名
        column_data: 列数据列表，长度必须与行数匹配
        default_value: 默认值，当column_data为None时使用
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 添加新列
        if column_data is not None:
            if len(column_data) != len(df):
                return False, f"列数据长度({len(column_data)})与表格行数({len(df)})不匹配"
            df[column_name] = column_data
        else:
            df[column_name] = default_value
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功添加列 '{column_name}' 到 {sheet_name}"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"添加列失败: {str(e)}"


def excel_delete_row(file_path: str, sheet_name: str, condition: dict = None, 
                     row_indices: list = None, save_path: str = None) -> Tuple[bool, str]:
    """
    删除Excel文件中的行
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        condition: 删除条件，字典格式 {列名: 值}，删除匹配的行
        row_indices: 要删除的行索引列表
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        original_len = len(df)
        
        # 根据条件删除行
        if condition:
            for col, val in condition.items():
                if col not in df.columns:
                    return False, f"列 '{col}' 不存在"
                df = df[df[col] != val]
        
        # 根据索引删除行
        if row_indices:
            df = df.drop(row_indices, errors='ignore')
        
        deleted_count = original_len - len(df)
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功删除 {deleted_count} 行"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"删除行失败: {str(e)}"


def excel_delete_column(file_path: str, sheet_name: str, column_names: Union[str, List[str]], 
                        save_path: str = None) -> Tuple[bool, str]:
    """
    删除Excel文件中的列
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_names: 要删除的列名或列名列表
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 确保column_names是列表
        if isinstance(column_names, str):
            column_names = [column_names]
        
        # 检查列是否存在
        missing_cols = [col for col in column_names if col not in df.columns]
        if missing_cols:
            return False, f"列不存在: {', '.join(missing_cols)}"
        
        # 删除列
        df = df.drop(columns=column_names)
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功删除列: {', '.join(column_names)}"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"删除列失败: {str(e)}"


def excel_delete_empty_rows(file_path: str, sheet_name: str, save_path: str = None) -> Tuple[bool, str]:
    """
    删除Excel文件中的空行
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        original_len = len(df)
        
        # 删除所有列都为空的行
        df = df.dropna(how='all')
        
        deleted_count = original_len - len(df)
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功删除 {deleted_count} 个空行"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"删除空行失败: {str(e)}"


def excel_modify_cell(file_path: str, sheet_name: str, row_index: int, 
                     column_name: str, new_value, save_path: str = None) -> Tuple[bool, str]:
    """
    修改Excel文件中的单元格值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        row_index: 行索引（从0开始）
        column_name: 列名
        new_value: 新值
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 检查行索引
        if row_index >= len(df) or row_index < 0:
            return False, f"行索引 {row_index} 超出范围 (0-{len(df)-1})"
        
        # 检查列名
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        # 修改单元格
        old_value = df.at[row_index, column_name]
        df.at[row_index, column_name] = new_value
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功修改单元格 [{row_index}, '{column_name}']: {old_value} -> {new_value}"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"修改单元格失败: {str(e)}"


def excel_modify_column(file_path: str, sheet_name: str, column_name: str, 
                       condition: dict, new_value, save_path: str = None) -> Tuple[bool, str]:
    """
    批量修改Excel文件中符合条件的列值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 要修改的列名
        condition: 条件字典 {列名: 值}
        new_value: 新值
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 检查列名
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        # 应用条件
        mask = pd.Series([True] * len(df))
        for col, val in condition.items():
            if col not in df.columns:
                return False, f"条件列 '{col}' 不存在"
            mask &= (df[col] == val)
        
        # 修改值
        modified_count = mask.sum()
        df.loc[mask, column_name] = new_value
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"成功修改 {modified_count} 个单元格"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"批量修改失败: {str(e)}"
