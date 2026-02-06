"""
Excel query and filter operations module
数据查询与筛选函数
"""

import pandas as pd
import os
from typing import Tuple, Union, List, Dict


def excel_query_data(file_path: str, sheet_name: str, condition: Dict = None, 
                     columns: List[str] = None) -> Tuple[bool, Union[pd.DataFrame, str]]:
    """
    按条件查询Excel数据
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        condition: 查询条件字典 {列名: 值} 或 {列名: (操作符, 值)}
        columns: 要返回的列名列表，None表示返回所有列
        
    Returns:
        (成功/失败, DataFrame或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 应用查询条件
        if condition:
            mask = pd.Series([True] * len(df))
            for col, val in condition.items():
                if col not in df.columns:
                    return False, f"列 '{col}' 不存在"
                
                # 支持简单值匹配和元组操作符
                if isinstance(val, tuple) and len(val) == 2:
                    op, value = val
                    if op == '>':
                        mask &= (df[col] > value)
                    elif op == '<':
                        mask &= (df[col] < value)
                    elif op == '>=':
                        mask &= (df[col] >= value)
                    elif op == '<=':
                        mask &= (df[col] <= value)
                    elif op == '!=':
                        mask &= (df[col] != value)
                    elif op == '==':
                        mask &= (df[col] == value)
                    elif op == 'in':
                        mask &= df[col].isin(value)
                    elif op == 'contains':
                        mask &= df[col].astype(str).str.contains(str(value), na=False)
                    else:
                        return False, f"不支持的操作符: {op}"
                else:
                    mask &= (df[col] == val)
            
            df = df[mask]
        
        # 选择指定列
        if columns:
            missing_cols = [col for col in columns if col not in df.columns]
            if missing_cols:
                return False, f"列不存在: {', '.join(missing_cols)}"
            df = df[columns]
        
        return True, df
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"查询数据失败: {str(e)}"


def excel_filter_by_value(file_path: str, sheet_name: str, column_name: str, 
                          values: List, save_path: str = None) -> Tuple[bool, str]:
    """
    按值筛选Excel数据并保存
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 筛选列名
        values: 要保留的值列表
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        original_len = len(df)
        df = df[df[column_name].isin(values)]
        filtered_len = len(df)
        
        if save_path:
            df.to_excel(save_path, sheet_name=sheet_name, index=False)
            return True, f"筛选完成: 保留 {filtered_len}/{original_len} 行"
        else:
            return True, df
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"筛选数据失败: {str(e)}"


def excel_search_text(file_path: str, sheet_name: str, search_text: str, 
                     columns: List[str] = None) -> Tuple[bool, Union[pd.DataFrame, str]]:
    """
    在Excel中搜索文本
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        search_text: 搜索文本
        columns: 要搜索的列名列表，None表示搜索所有列
        
    Returns:
        (成功/失败, 包含搜索结果的DataFrame或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 选择要搜索的列
        search_cols = columns if columns else df.columns.tolist()
        
        # 检查列是否存在
        missing_cols = [col for col in search_cols if col not in df.columns]
        if missing_cols:
            return False, f"列不存在: {', '.join(missing_cols)}"
        
        # 搜索文本
        mask = pd.Series([False] * len(df))
        for col in search_cols:
            mask |= df[col].astype(str).str.contains(search_text, case=False, na=False)
        
        result_df = df[mask]
        
        return True, result_df
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"搜索文本失败: {str(e)}"


def excel_get_unique_values(file_path: str, sheet_name: str, column_name: str) -> Tuple[bool, Union[List, str]]:
    """
    获取Excel列中的唯一值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名
        
    Returns:
        (成功/失败, 唯一值列表或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        unique_values = df[column_name].unique().tolist()
        
        return True, unique_values
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"获取唯一值失败: {str(e)}"


def excel_filter_by_range(file_path: str, sheet_name: str, column_name: str, 
                         min_value=None, max_value=None, save_path: str = None) -> Tuple[bool, Union[str, pd.DataFrame]]:
    """
    按数值范围筛选Excel数据
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 筛选列名
        min_value: 最小值（包含）
        max_value: 最大值（包含）
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息或DataFrame)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        original_len = len(df)
        
        # 应用范围筛选
        if min_value is not None:
            df = df[df[column_name] >= min_value]
        if max_value is not None:
            df = df[df[column_name] <= max_value]
        
        filtered_len = len(df)
        
        if save_path:
            df.to_excel(save_path, sheet_name=sheet_name, index=False)
            return True, f"筛选完成: 保留 {filtered_len}/{original_len} 行"
        else:
            return True, df
        
    except ValueError as e:
        return False, f"工作表名错误或数据类型错误: {str(e)}"
    except Exception as e:
        return False, f"范围筛选失败: {str(e)}"
