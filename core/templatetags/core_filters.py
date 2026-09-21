import json
from django import template

register = template.Library()

@register.filter
def tojson(value):
    return json.dumps(value, ensure_ascii=False)

@register.filter
def get_item(dictionary, key):
    """Akses dict dengan key yang mengandung karakter khusus seperti tanda -"""
    if isinstance(dictionary, dict):
        return dictionary.get(key, '')
    return ''
