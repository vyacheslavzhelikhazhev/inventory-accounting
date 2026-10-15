"""Fill the supplied handover form without carrying the sample's personal data."""
import copy
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from enhancements import paragraphs,replace_text,validate_markers
from prices import money,total

REPRESENTATIVE=''
AUTHORITY=''

def short_name(name):
    parts=name.split()
    return parts[0]+' '+''.join(x[0]+'.' for x in parts[1:]) if len(parts)>1 else name

def fill_handover(d,number,date,operation,target,items,p,settings):
    employee=target
    shared=p.get('target_kind')=='Объект'
    if shared:
        group=[e for e in items if e['action']=='Выдача'] or items
        names={(e['after'] if e['action']=='Выдача' else e['before']).get('responsible','').strip() for e in group}
        employee=next(iter(names)) if len(names)==1 and '' not in names else '________________________'
    authority=settings.get('authority',AUTHORITY).strip() or '________________________________'
    representative=settings.get('employer_representative','').strip() or '________________________________'
    employer_role=settings.get('employer_role','').strip() or '________________________________'
    parties=f'Работодатель (представитель: {employer_role}, {representative}), действующий на основании документа «{authority}», с одной стороны, и {employee}, проживающий по адресу: __________________________________________________, паспорт: ______ № __________, выдан ______________________________________________ «___» __________ 20___ года, именуемый в дальнейшем «Работник», с другой стороны, составили настоящий акт о следующем:'
    if operation=='Выдача':description='Работодатель передал, а Работник принял следующие материальные ценности для выполнения своих должностных обязанностей:'
    elif operation in ['Возврат','Увольнение']:description='Работник передал, а Работодатель принял на склад следующие материальные ценности:'
    else:description='При замене оборудования Работник вернул, а Работодатель выдал взамен следующие материальные ценности. Направление передачи указано в каждой строке:'
    if shared:description+=' Место общего использования: '+target+'.'
    def summary(entries,label):
        amount=total(entries) or '________________'
        return f'{label}: {len(entries)} шт., на сумму {amount} руб.'
    if operation=='Замена':
        summary_text=summary([e for e in items if e['action']=='Возврат'],'Возвращено')+'\n'+summary([e for e in items if e['action']=='Выдача'],'Выдано взамен')
    else:summary_text=summary(items,'Итого принято материальных ценностей')
    values={'NUMBER':number,'DATE':date,'PARTIES':parties,'DESCRIPTION':description,'SUMMARY':summary_text,'CHECKED':'Материальные ценности проверены сторонами при передаче.','OPERATOR_SHORT':short_name(representative),'EMPLOYEE_SHORT':short_name(employee)}
    for par in list(paragraphs(d)):
        for key,value in values.items():
            token='{{'+key+'}}'
            if token in par.text:replace_text(par,token,value)
    candidates=[t for t in d.tables if len(t.columns)==5 and 'Стоимость' in t.rows[0].cells[-1].text]
    if len(candidates)!=1 or len(candidates[0].rows)<2:raise ValueError('В шаблоне приема-передачи нужна таблица из пяти столбцов и строка-образец оборудования.')
    table=candidates[0];prototype=copy.deepcopy(table.rows[1]._tr)
    while len(table.rows)>1:table._tbl.remove(table.rows[-1]._tr)
    for i,entry in enumerate(items,1):
        a=entry['after'];table._tbl.append(copy.deepcopy(prototype));row=table.rows[-1]
        name=a['name']
        if operation=='Замена':name=('Возврат: ' if entry['action']=='Возврат' else 'Выдача: ')+name
        if a.get('kit'):name+='\nКомплектность: '+a['kit']
        if a.get('condition'):name+='\nСостояние: '+a['condition']
        if entry.get('note'):name+='\n'+entry['note']
        serial=a.get('serial') or 'ID '+str(a['id'])
        for index,(cell,value) in enumerate(zip(row.cells,[str(i)+'.',name,serial,'1',money(a.get('purchase_price'))])):
            cell.text=value
            for par in cell.paragraphs:
                par.alignment=WD_ALIGN_PARAGRAPH.LEFT if index==1 else WD_ALIGN_PARAGRAPH.CENTER
                par.paragraph_format.space_before=Pt(4);par.paragraph_format.space_after=Pt(4)
                for run in par.runs:run.font.size=Pt(10);run.bold=False
        pr=row._tr.get_or_add_trPr()
        for height in list(pr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trHeight')):pr.remove(height)
        pr.append(OxmlElement('w:cantSplit'))
    validate_markers(d)
