"""
Excel Functions Package
Excel处理基础函数包

所有函数统一导出，支持外部直接调用
"""

# CRUD operations (增删改查)
from .crud import (
    excel_add_row,
    excel_add_column,
    excel_delete_row,
    excel_delete_column,
    excel_delete_empty_rows,
    excel_modify_cell,
    excel_modify_column
)

# Query operations (查询筛选)
from .query import (
    excel_query_data,
    excel_filter_by_value,
    excel_search_text,
    excel_get_unique_values,
    excel_filter_by_range
)

# Statistics operations (统计计算)
from .statistics import (
    excel_sum_column,
    excel_average_column,
    excel_count_values,
    excel_max_value,
    excel_min_value,
    excel_deduplicate,
    excel_group_statistics,
    excel_calculate_statistics
)

# Merge operations (合并操作)
from .merge import (
    excel_merge_files,
    excel_merge_sheets,
    excel_join_files,
    excel_append_data
)

__all__ = [
    # CRUD
    'excel_add_row',
    'excel_add_column',
    'excel_delete_row',
    'excel_delete_column',
    'excel_delete_empty_rows',
    'excel_modify_cell',
    'excel_modify_column',
    # Query
    'excel_query_data',
    'excel_filter_by_value',
    'excel_search_text',
    'excel_get_unique_values',
    'excel_filter_by_range',
    # Statistics
    'excel_sum_column',
    'excel_average_column',
    'excel_count_values',
    'excel_max_value',
    'excel_min_value',
    'excel_deduplicate',
    'excel_group_statistics',
    'excel_calculate_statistics',
    # Merge
    'excel_merge_files',
    'excel_merge_sheets',
    'excel_join_files',
    'excel_append_data',
]
