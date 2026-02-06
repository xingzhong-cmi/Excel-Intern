"""
Excel statistics and calculation operations module
数据统计与计算函数
"""

import pandas as pd
import os
from typing import Tuple, Union, Dict


def excel_sum_column(file_path: str, sheet_name: str, column_name: str) -> Tuple[bool, Union[float, str]]:
    """
    计算Excel列的求和
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名
        
    Returns:
        (成功/失败, 求和结果或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        total = df[column_name].sum()
        
        return True, total
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"求和失败: {str(e)}"


def excel_average_column(file_path: str, sheet_name: str, column_name: str) -> Tuple[bool, Union[float, str]]:
    """
    计算Excel列的平均值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名
        
    Returns:
        (成功/失败, 平均值或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        avg = df[column_name].mean()
        
        return True, avg
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"计算平均值失败: {str(e)}"


def excel_count_values(file_path: str, sheet_name: str, column_name: str = None) -> Tuple[bool, Union[int, str]]:
    """
    计数Excel中的非空值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名，None表示统计总行数
        
    Returns:
        (成功/失败, 计数结果或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name is None:
            count = len(df)
        else:
            if column_name not in df.columns:
                return False, f"列 '{column_name}' 不存在"
            count = df[column_name].count()
        
        return True, count
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"计数失败: {str(e)}"


def excel_max_value(file_path: str, sheet_name: str, column_name: str) -> Tuple[bool, Union[float, str]]:
    """
    获取Excel列的最大值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名
        
    Returns:
        (成功/失败, 最大值或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        max_val = df[column_name].max()
        
        return True, max_val
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"获取最大值失败: {str(e)}"


def excel_min_value(file_path: str, sheet_name: str, column_name: str) -> Tuple[bool, Union[float, str]]:
    """
    获取Excel列的最小值
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名
        
    Returns:
        (成功/失败, 最小值或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        min_val = df[column_name].min()
        
        return True, min_val
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"获取最小值失败: {str(e)}"


def excel_deduplicate(file_path: str, sheet_name: str, columns: list = None, 
                     keep: str = 'first', save_path: str = None) -> Tuple[bool, str]:
    """
    Excel数据去重
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        columns: 用于判断重复的列名列表，None表示所有列
        keep: 保留策略 'first'/'last'/False
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        original_len = len(df)
        
        # 检查列是否存在
        if columns:
            missing_cols = [col for col in columns if col not in df.columns]
            if missing_cols:
                return False, f"列不存在: {', '.join(missing_cols)}"
        
        # 去重
        df = df.drop_duplicates(subset=columns, keep=keep)
        duplicates_removed = original_len - len(df)
        
        output_path = save_path if save_path else file_path
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        
        return True, f"去重完成: 删除 {duplicates_removed} 个重复行，保留 {len(df)} 行"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"去重失败: {str(e)}"


def excel_group_statistics(file_path: str, sheet_name: str, group_by_column: str, 
                           agg_column: str, agg_func: str = 'sum') -> Tuple[bool, Union[pd.DataFrame, str]]:
    """
    分组统计Excel数据
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        group_by_column: 分组列名
        agg_column: 聚合列名
        agg_func: 聚合函数 'sum'/'mean'/'count'/'max'/'min'
        
    Returns:
        (成功/失败, 统计结果DataFrame或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if group_by_column not in df.columns:
            return False, f"分组列 '{group_by_column}' 不存在"
        if agg_column not in df.columns:
            return False, f"聚合列 '{agg_column}' 不存在"
        
        # 分组聚合
        agg_funcs = {
            'sum': 'sum',
            'mean': 'mean',
            'average': 'mean',
            'count': 'count',
            'max': 'max',
            'min': 'min'
        }
        
        if agg_func not in agg_funcs:
            return False, f"不支持的聚合函数: {agg_func}"
        
        result = df.groupby(group_by_column)[agg_column].agg(agg_funcs[agg_func]).reset_index()
        
        return True, result
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"分组统计失败: {str(e)}"


def excel_calculate_statistics(file_path: str, sheet_name: str, column_name: str) -> Tuple[bool, Union[Dict, str]]:
    """
    计算Excel列的综合统计信息
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 列名
        
    Returns:
        (成功/失败, 统计信息字典或错误消息)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            return False, f"列 '{column_name}' 不存在"
        
        stats = {
            '总数': df[column_name].count(),
            '求和': df[column_name].sum() if pd.api.types.is_numeric_dtype(df[column_name]) else 'N/A',
            '平均值': df[column_name].mean() if pd.api.types.is_numeric_dtype(df[column_name]) else 'N/A',
            '最大值': df[column_name].max(),
            '最小值': df[column_name].min(),
            '标准差': df[column_name].std() if pd.api.types.is_numeric_dtype(df[column_name]) else 'N/A',
            '中位数': df[column_name].median() if pd.api.types.is_numeric_dtype(df[column_name]) else 'N/A',
            '唯一值数量': df[column_name].nunique(),
        }
        
        return True, stats
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"计算统计信息失败: {str(e)}"
