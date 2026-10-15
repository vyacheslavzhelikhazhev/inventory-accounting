"""Purchase prices: empty is unknown; Decimal strings preserve kopecks."""
from decimal import Decimal, InvalidOperation
import re

def price(value):
    if value is None or value=='':return ''
    text=str(value).strip().replace('\u00a0','').replace(' ','').replace(',','.')
    if not text:return ''
    if not re.fullmatch(r'\d+(?:\.\d{1,2})?',text):
        raise ValueError('Стоимость: укажите неотрицательную сумму с точностью до копеек или оставьте поле пустым.')
    amount=Decimal(text)
    if amount>Decimal('999999999.99'):raise ValueError('Стоимость превышает допустимое значение 999 999 999,99 руб.')
    return format(amount,'.2f')

def money(value):
    val=price(value)
    return format(Decimal(val),',.2f').replace(',',' ').replace('.',',') if val else ''

def total(entries):
    values=[price(e['after'].get('purchase_price')) for e in entries]
    if not values or any(x=='' for x in values):return ''
    return money(sum((Decimal(x) for x in values),Decimal('0')))
