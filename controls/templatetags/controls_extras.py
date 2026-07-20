from django import template
from django.utils.safestring import mark_safe

register = template.Library()


@register.filter
def get_item(dictionary, key):
    if not dictionary:
        return ""
    return dictionary.get(key, "")


@register.filter
def word_diff(old_text, new_text):
    from controls.services import compute_word_diff_html
    return mark_safe(compute_word_diff_html(old_text, new_text))
