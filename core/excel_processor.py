# core/excel_processor.py
# (代码与第25轮回复中的版本完全相同，此处不再重复)
# 请确保您使用的是那个版本，它正确处理了K值的初始化。
import pandas as pd
from typing import List, Dict, Tuple, Any
import os
import logging

logger = logging.getLogger(__name__)


def read_excel_files(file_paths: List[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    logger.info(f"read_excel_files: 开始处理 {len(file_paths)} 个Excel文件。")
    all_parameters_raw: List[Dict[str, Any]] = []
    errors: List[str] = []

    for file_idx, file_path in enumerate(file_paths):
        logger.debug(f"  正在处理文件 {file_idx + 1}/{len(file_paths)}: {file_path}")
        try:
            engine = None
            if file_path.lower().endswith('.xlsx'):
                engine = 'openpyxl'
            elif file_path.lower().endswith('.xls'):
                engine = 'xlrd'
            else:
                errors.append(f"不支持的文件类型: {os.path.basename(file_path)}。仅支持 .xls 和 .xlsx。")
                continue

            df = pd.read_excel(file_path, header=None, sheet_name=0, engine=engine)
            if df.shape[0] < 13:
                errors.append(f"文件 '{os.path.basename(file_path)}' 的行数少于14行，无法处理。")
                continue

            col_count = df.shape[1]

            for row_idx, row in df.iloc[13:].iterrows():
                excel_row_display = row_idx + 14
                param_name = str(row.iloc[0]).strip() if col_count > 0 and pd.notna(row.iloc[0]) else ""
                if not param_name:
                    logger.debug(
                        f"    文件 '{os.path.basename(file_path)}' Excel行 {excel_row_display} 参数名为空，跳过。")
                    continue

                nominal_value = str(row.iloc[2]).strip() if col_count > 2 and pd.notna(row.iloc[2]) else ""
                upper_tol_str = str(row.iloc[3]).strip() if col_count > 3 and pd.notna(row.iloc[3]) else ""
                lower_tol_str = str(row.iloc[4]).strip() if col_count > 4 and pd.notna(row.iloc[4]) else ""

                k2003_val = ""
                k2005_val = "0"
                k2009_val = "0"
                k2142_val = ""  # 修正2：K2142 始终默认空

                is_upper_tol_present_in_excel = (upper_tol_str.strip() != "")
                is_lower_tol_present_in_excel = (lower_tol_str.strip() != "")

                k2121_initial_val = '1' if is_upper_tol_present_in_excel else '0'
                k2120_initial_val = '1' if is_lower_tol_present_in_excel else '0'

                all_parameters_raw.append({
                    "K2001_val": param_name, "K2002_val": param_name,
                    "K2101_val": nominal_value,
                    "K2113_val": upper_tol_str, "K2112_val": lower_tol_str,
                    "K2142_val": k2142_val, "K2003_val": k2003_val,
                    "K2005_val": k2005_val, "K2009_val": k2009_val,
                    "K2121_val": k2121_initial_val, "K2120_val": k2120_initial_val,
                    "selected_for_output": True,
                    "source_file": os.path.basename(file_path),
                    "original_row_index_df": row_idx,
                    "original_excel_row": excel_row_display
                })
        except Exception as e:
            errors.append(f"处理文件 '{os.path.basename(file_path)}' 时出错: {e}")
            logger.critical(f"    处理文件 '{os.path.basename(file_path)}' 时发生严重错误: {e}", exc_info=True)

    deduplicated_parameters: List[Dict[str, Any]] = []
    seen_params = set()
    for param in all_parameters_raw:
        param_key = (param["K2001_val"], param["K2002_val"])
        if param_key not in seen_params:
            deduplicated_parameters.append(param)
            seen_params.add(param_key)
    if not deduplicated_parameters and not errors and file_paths:
        errors.append("在所有选择的Excel文件中，从第14行开始未找到有效的参数数据，或者所有参数名为空。")
    return deduplicated_parameters, errors