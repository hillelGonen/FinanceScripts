import json
import os

DEFAULT_CATEGORY_MAPPING = {
    "Groceries & Supermarket": ["רמי לוי", "שופרסל", "AM:PM"],
    "Uncategorized": []
}


def load_category_mapping(path='categories.json'):
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return {k: v for k, v in data.items() if k != 'budgets'}
    except Exception:
        pass
    return DEFAULT_CATEGORY_MAPPING


def load_budgets(path='categories.json'):
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('budgets', {})
    except Exception:
        pass
    return {}


def save_categories_json(content_str, path='categories.json'):
    try:
        data = json.loads(content_str)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True, None
    except Exception as e:
        return False, str(e)


def classify_expense(merchant_name, mapping):
    if not isinstance(merchant_name, str):
        return "Uncategorized"
    merchant_upper = merchant_name.upper()
    for category, keywords in mapping.items():
        for keyword in keywords:
            if keyword.upper() in merchant_upper:
                return category
    return "Uncategorized"
