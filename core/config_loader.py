import json
import os

CONFIG_PATH = os.path.join(os.getcwd(), 'config', 'groups.json')

def load_group_config():
    """加载默认分组配置，如果文件不存在则返回 None"""
    if not os.path.exists(CONFIG_PATH):
        return None
    
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading config: {e}")
        return None

def parse_group_config(config_data):
    """
    将配置字典转换为 processor 需要的两个查找表：
    1. group_summaries: {id_str: name}
    2. item_to_group_id: {item_name: id_int}
    """
    if not config_data:
        return {}, {}

    group_summaries = {}
    item_to_group_id = {}

    for group in config_data.get('groups', []):
        gid_str = group['id']
        try:
            gid_int = int(gid_str)
        except ValueError:
            continue
        
        group_summaries[gid_str] = group['name']
        
        for item in group.get('items', []):
            if item not in item_to_group_id:
                item_to_group_id[item] = []
            item_to_group_id[item].append(gid_int)
            
    return group_summaries, item_to_group_id

def get_processor_config():
    """加载并解析默认配置"""
    config = load_group_config()
    return parse_group_config(config)
