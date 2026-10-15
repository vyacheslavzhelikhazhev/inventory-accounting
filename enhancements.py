"""Portable document helpers shared by the application and regression tests."""
import re

def paragraphs(document):
    seen=set()
    def walk(container):
        for p in container.paragraphs:
            if p._element not in seen:
                seen.add(p._element);yield p
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:yield from walk(cell)
    yield from walk(document)
    for section in document.sections:
        for part in [section.header,section.footer,section.first_page_header,section.first_page_footer,section.even_page_header,section.even_page_footer]:
            yield from walk(part)

def replace_text(par,old,new):
    """Replace across runs while retaining formatting outside the replaced marker."""
    starts=[m.start() for m in re.finditer(re.escape(old),par.text)]
    for start in reversed(starts):
        end=start+len(old);pos=0;placed=False
        for run in par.runs:
            text=run.text;lo=pos;hi=pos+len(text);pos=hi
            if hi<=start or lo>=end:continue
            left=text[:max(0,start-lo)];right=text[max(0,end-lo):]
            run.text=left+(new if not placed else '')+right;placed=True
        if not placed:
            raise ValueError('Шаблон содержит неподдерживаемое поле. Замените его обычным текстовым маркером.')

def without_organization(document):
    for p in list(paragraphs(document)):
        if '{{ORG}}' in p.text or p.text.strip().lower().startswith('организация:'):
            # Leave empty cells valid; deleting all their paragraphs damages Word XML.
            if p._element.getparent().tag.endswith('}tc'):
                for r in p.runs:r.text=''
            else:p._element.getparent().remove(p._element)
        elif 'Представитель организации:' in p.text:
            replace_text(p,'Представитель организации:','Выдал / принял:')

def equipment_table(document):
    candidates=[]
    for table in document.tables:
        if not table.rows:continue
        headers=[c.text.strip().casefold() for c in table.rows[0].cells]
        if len(headers)==6 and headers[0]=='id' and 'действие' in headers[1] and 'оборудование' in headers[2] and 'серийн' in headers[3]:
            candidates.append(table)
    if len(candidates)!=1:
        raise ValueError('В шаблоне нужна одна таблица оборудования с 6 столбцами: ID, Действие, Оборудование, Серийный номер, Комплектность, Состояние и замечания.')
    return candidates[0]

def validate_markers(document):
    markers=sorted(set(re.findall(r'\{\{[^{}]+\}\}','\n'.join(p.text for p in paragraphs(document)))))
    if markers:raise ValueError('В шаблоне остались незаполненные поля: '+', '.join(markers))
