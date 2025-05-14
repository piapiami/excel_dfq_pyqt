# core/config_manager.py
import json
import os
from typing import List, Dict, Any

# 定义 config.json 的基本名称
BASE_CONFIG_FILENAME = "config.json"

# 获取 core 目录的父目录，即项目根目录
# __file__ 是当前脚本 (config_manager.py) 的路径
# os.path.dirname(__file__) 是 core 目录
# os.path.dirname(os.path.dirname(__file__)) 是项目根目录
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 构建 config.json 的绝对路径
CONFIG_FILE_PATH_ABSOLUTE = os.path.join(PROJECT_ROOT, BASE_CONFIG_FILENAME)

DEFAULT_CONFIG = {
    "OutputPath": "",
    "SystemSettings": [
        {
            "K1001": "DefaultPartNumber",
            "K1002": "DefaultPartName",
            "K1086": "DefaultStation",
            "K1091": "DefaultLine",
            "K1004": "5"  # 默认SPC送检数量
        }
    ]
}


def load_config() -> Dict[str, Any]:
    """从绝对路径加载 config.json。"""
    if not os.path.exists(CONFIG_FILE_PATH_ABSOLUTE):
        save_config(DEFAULT_CONFIG)  # save_config 现在使用绝对路径
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_FILE_PATH_ABSOLUTE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            # 确保所有必要的键都存在，如果不存在则使用默认值
            if "OutputPath" not in config:
                config["OutputPath"] = DEFAULT_CONFIG["OutputPath"]
            if "SystemSettings" not in config or not isinstance(config.get("SystemSettings"), list):
                config["SystemSettings"] = DEFAULT_CONFIG["SystemSettings"]
            else:
                default_k1004 = DEFAULT_CONFIG["SystemSettings"][0].get("K1004", "5")
                for setting in config["SystemSettings"]:
                    if not isinstance(setting, dict):  # 如果列表中的项不是字典，则跳过或替换
                        # 为了健壮性，可以考虑如何处理错误格式的条目
                        continue
                    if "K1001" not in setting: setting["K1001"] = ""  # 提供默认空字符串
                    if "K1002" not in setting: setting["K1002"] = ""
                    if "K1086" not in setting: setting["K1086"] = ""
                    if "K1091" not in setting: setting["K1091"] = ""
                    if "K1004" not in setting or not setting.get("K1004"):  # 检查K1004是否存在或为空
                        setting["K1004"] = default_k1004
            return config
    except (json.JSONDecodeError, IOError) as e:
        print(f"加载配置文件 {CONFIG_FILE_PATH_ABSOLUTE} 出错: {e}. 将重置为默认配置。")
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()


def save_config(config: Dict[str, Any]):
    """将配置保存到绝对路径的 config.json。"""
    try:
        with open(CONFIG_FILE_PATH_ABSOLUTE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except IOError as e:
        print(f"保存配置文件到 {CONFIG_FILE_PATH_ABSOLUTE} 失败: {e}")


def get_system_settings() -> List[Dict[str, str]]:
    """获取系统设置列表。"""
    config = load_config()
    # 确保返回的是一个列表，即使配置中 SystemSettings 格式错误或丢失
    settings = config.get("SystemSettings", [])
    return settings if isinstance(settings, list) else DEFAULT_CONFIG["SystemSettings"]


def get_output_path() -> str:
    """获取输出路径。"""
    config = load_config()
    path = config.get("OutputPath", "")
    return path if isinstance(path, str) else DEFAULT_CONFIG["OutputPath"]


def update_system_settings(settings: List[Dict[str, str]]):
    """更新系统设置并保存。"""
    config = load_config()  # 加载当前配置，确保其他部分不受影响

    # 为传入的settings中的每个条目确保K1004字段存在且有值
    default_k1004_value = DEFAULT_CONFIG["SystemSettings"][0].get("K1004", "5")
    processed_settings = []
    if isinstance(settings, list):
        for setting in settings:
            if isinstance(setting, dict):
                # 确保所有预期的K值键都存在，如果不存在，可以添加空字符串或默认值
                setting.setdefault("K1001", "")
                setting.setdefault("K1002", "")
                setting.setdefault("K1086", "")
                setting.setdefault("K1091", "")
                if "K1004" not in setting or not setting.get("K1004"):
                    setting["K1004"] = default_k1004_value
                processed_settings.append(setting)

    config["SystemSettings"] = processed_settings
    save_config(config)


def update_output_path(path: str):
    """更新输出路径并保存。"""
    config = load_config()
    config["OutputPath"] = path if isinstance(path, str) else DEFAULT_CONFIG["OutputPath"]
    save_config(config)