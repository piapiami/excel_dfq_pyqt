# core/dfq_writer.py
import logging
from typing import List, Dict, Any, Tuple
import datetime
import os

logger = logging.getLogger(__name__)


def generate_dfq_content(parameters_to_output: List[Dict[str, Any]], header_info: Dict[str, str]) -> List[str]:
    logger.info(f"generate_dfq_content: 开始生成DFQ内容，参数数量: {len(parameters_to_output)}")
    logger.debug(f"  抬头信息: {header_info}")

    dfq_lines: List[str] = []
    param_count = len(parameters_to_output)

    dfq_lines.append(f"K0100 {param_count}")
    dfq_lines.append(f"K1001 {header_info.get('K1001', '')}")
    dfq_lines.append(f"K1002 {header_info.get('K1002', '')}")
    dfq_lines.append(f"K1004 {header_info.get('K1004', '5')}")
    dfq_lines.append(f"K1086 {header_info.get('K1086', '')}")
    dfq_lines.append(f"K1091 {header_info.get('K1091', '')}")

    for i, param_data in enumerate(parameters_to_output):
        param_index_output = i + 1
        k2001_val = param_data.get('K2001_val', '')
        k2009_val = param_data.get('K2009_val', '0')  # 获取K2009的值
        k2120_val = param_data.get('K2120_val', '0')
        k2121_val = param_data.get('K2121_val', '0')
        logger.debug(
            f"  正在写入参数 {param_index_output}: {k2001_val}, K2009='{k2009_val}', K2120='{k2120_val}', K2121='{k2121_val}'")

        dfq_lines.append(f"K2001/{param_index_output} {k2001_val}")
        dfq_lines.append(f"K2002/{param_index_output} {param_data.get('K2002_val', '')}")
        dfq_lines.append(f"K2003/{param_index_output} {param_data.get('K2003_val', '')}")
        dfq_lines.append(f"K2005/{param_index_output} {param_data.get('K2005_val', '0')}")
        dfq_lines.append(f"K2009/{param_index_output} {k2009_val}")  # 使用获取到的K2009值
        dfq_lines.append(f"K2101/{param_index_output} {param_data.get('K2101_val', '')}")
        dfq_lines.append(f"K2113/{param_index_output} {param_data.get('K2113_val', '')}")
        dfq_lines.append(f"K2112/{param_index_output} {param_data.get('K2112_val', '')}")
        dfq_lines.append(f"K2121/{param_index_output} {k2121_val}")  # 使用获取到的K2121值
        dfq_lines.append(f"K2120/{param_index_output} {k2120_val}")  # 使用获取到的K2120值
        dfq_lines.append(f"K2142/{param_index_output} {param_data.get('K2142_val', '')}")

    logger.info("DFQ内容生成完毕。")
    return dfq_lines


def write_dfq_file(output_path: str, dfq_lines: List[str], header_info: Dict[str, str]) -> Tuple[bool, str]:
    # ... (此函数与上一轮代码一致)
    logger.info(f"write_dfq_file: 准备写入DFQ文件到路径: {output_path}")
    if not os.path.isdir(output_path):
        try:
            os.makedirs(output_path, exist_ok=True)
            logger.info(f"输出目录 '{output_path}' 不存在，已创建。")
        except OSError as e:
            logger.error(f"创建输出目录 '{output_path}' 失败: {e}", exc_info=True)
            return False, f"创建输出目录 '{output_path}' 失败: {e}"

    k1001 = header_info.get('K1001', 'NA')
    k1002 = header_info.get('K1002', 'NA')
    k1086 = header_info.get('K1086', 'NA')
    k1091 = header_info.get('K1091', 'NA')
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    def sanitize(name):
        return "".join(c if c.isalnum() or c in ('_', '-') else '_' for c in name)

    filename = f"{sanitize(k1001)}_{sanitize(k1002)}_{sanitize(k1086)}_{sanitize(k1091)}_{timestamp}.dfq"
    file_path = os.path.join(output_path, filename)
    logger.debug(f"目标DFQ文件名: {file_path}")

    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            for line in dfq_lines:
                f.write(line + "\n")
        logger.info(f"DFQ文件成功写入到: {file_path}")
        return True, file_path
    except IOError as e:
        logger.error(f"写入 DFQ 文件 '{file_path}' 失败: {e}", exc_info=True)
        return False, f"写入 DFQ 文件 '{file_path}' 失败: {e}"