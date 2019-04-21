from docx import Document
from bcompiler.utils import project_data_from_master
from collections import OrderedDict
import datetime
from docx.oxml.ns import nsdecls
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.oxml import OxmlElement
from docx.shared import Cm, Inches, Pt, RGBColor
import difflib


def get_project_names(data):
    project_name_list = []
    for x in data:
        project_name_list.append(x)
    return project_name_list


def converting_RAGs(rag):
    if rag == 'Green':
        return 'G'
    elif rag == 'Amber/Green':
        return 'A/G'
    elif rag == 'Amber':
        return 'A'
    elif rag == 'Amber/Red':
        return 'A/R'
    else:
        return 'R'


def cell_colouring(cell, colour):
    if colour == 'R':
        colour = parse_xml(r'<w:shd {} w:fill="cb1f00"/>'.format(nsdecls('w')))
    elif colour == 'A/R':
        colour = parse_xml(r'<w:shd {} w:fill="f97b31"/>'.format(nsdecls('w')))
    elif colour == 'A':
        colour = parse_xml(r'<w:shd {} w:fill="fce553"/>'.format(nsdecls('w')))
    elif colour == 'A/G':
        colour = parse_xml(r'<w:shd {} w:fill="a5b700"/>'.format(nsdecls('w')))
    elif colour == 'G':
        colour = parse_xml(r'<w:shd {} w:fill="17960c"/>'.format(nsdecls('w')))

    cell._tc.get_or_add_tcPr().append(colour)


'''function places text into doc highlighing all changes'''


def compare_text_showall(text_1, text_2, doc):
    comp = difflib.Differ()
    diff = list(comp.compare(text_2.split(), text_1.split()))
    new_text = diff
    y = doc.add_paragraph()

    for i in range(0, len(diff)):
        f = len(diff) - 1
        if i < f:
            a = i - 1
        else:
            a = i

        if diff[i][0:3] == '  |':
            j = i + 1
            if diff[i][0:3] and diff[a][0:3] == '  |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '+ |':
            if diff[i][0:3] and diff[a][0:3] == '+ |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '- |':
            pass
        elif diff[i][0:3] == '  -':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0:3] == '  •':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0] == '+':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.color.rgb = RGBColor(255, 0, 0)
        elif diff[i][0] == '-':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.strike = True
        elif diff[i][0] == '?':
            pass
        else:
            if diff[i] != '+ |':
                y.add_run(diff[i])

    return doc


'''function places text into doc highlighing new and old text'''


def compare_text_newandold(text_1, text_2, doc):
    comp = difflib.Differ()
    diff = list(comp.compare(text_2.split(), text_1.split()))
    new_text = diff
    y = doc.add_paragraph()

    for i in range(0, len(diff)):
        f = len(diff) - 1
        if i < f:
            a = i - 1
        else:
            a = i

        if diff[i][0:3] == '  |':
            j = i + 1
            if diff[i][0:3] and diff[a][0:3] == '  |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '+ |':
            if diff[i][0:3] and diff[a][0:3] == '+ |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '- |':
            pass
        elif diff[i][0:3] == '  -':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0:3] == '  •':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0] == '+':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.color.rgb = RGBColor(255, 0, 0)
        elif diff[i][0] == '-':
            pass
        elif diff[i][0] == '?':
            pass
        else:
            if diff[i] != '+ |':
                y.add_run(diff[i][1:])

    return doc


def printing(name, dictionary_1, dictionary_2, dictionary_3, dictionary_4):
    doc = Document()
    print(name)
    heading = str(name)
    name = str(name)
    # TODO: change heading font size
    # todo be able to change text sixe and font
    intro = doc.add_heading(str(heading), 0)
    intro.alignment = 1
    intro.bold = True

    '''key names and contact details'''
    # doc.add_paragraph()
    # y = doc.add_paragraph()
    # heading = 'Project Leadership'
    # y.add_run(str(heading)).bold = True

    # doc.add_paragraph()
    y = doc.add_paragraph()
    a = dictionary_1[name]['SRO Full Name']
    if a == None:
        a = 'TBC'
    else:
        a = a
        b = dictionary_1[name]['SRO Phone No.']
        if b == None:
            b = 'TBC'
        else:
            b = b

    y.add_run('SRO name:  ' + str(a) + ',   Tele:  ' + str(b))

    y = doc.add_paragraph()
    a = dictionary_1[name]['PD Full Name']
    if a == None:
        a = 'TBC'
    else:
        a = a
        b = dictionary_1[name]['PD Phone No.']
        if b == None:
            b = 'TBC'
        else:
            b = b

    y.add_run('PD name:  ' + str(a) + ',   Tele:  ' + str(b))

    '''DCA information'''
    # y = doc.add_paragraph()
    # heading = 'Project confidence trends'
    # y.add_run(str(heading)).bold = True

    '''Start of table with DCA confidence ratings'''
    # doc.add_paragraph()
    table1 = doc.add_table(rows=1, cols=5)
    table1.cell(0, 0).width = Cm(7)

    table1.cell(0, 1).text = 'This quarter'  # this needs to be changed each quarter
    table1.cell(0, 2).text = 'Q2 1819'
    table1.cell(0, 3).text = 'Q1 1819'
    table1.cell(0, 4).text = 'Q4 1718'

    '''setting row height - partially working'''
    # todo understand row height better
    row = table1.rows[0]
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(200))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

    '''SRO DCA ratings'''

    table2 = doc.add_table(rows=1, cols=5)
    table2.cell(0, 0).width = Cm(7)
    table2.cell(0, 0).text = 'SRO DCA'
    a = converting_RAGs(dictionary_1[name]['Departmental DCA'])
    table2.cell(0, 1).text = a
    cell_colouring(table2.cell(0, 1), a)

    try:
        a = converting_RAGs(dictionary_2[name]['Departmental DCA'])
        table2.cell(0, 2).text = a
        cell_colouring(table2.cell(0, 2), a)
    except KeyError:
        table2.cell(0, 2).text = 'N/A'

    try:
        a = converting_RAGs(dictionary_3[name]['Departmental DCA'])
        table2.cell(0, 3).text = a
        cell_colouring(table2.cell(0, 3), a)
    except KeyError:
        table2.cell(0, 3).text = 'N/A'

    try:
        a = converting_RAGs(dictionary_4[name]['Departmental DCA'])
        table2.cell(0, 4).text = a
        cell_colouring(table2.cell(0, 4), a)
    except KeyError:
        table2.cell(0, 4).text = 'N/A'

    '''SRO Financial confidence'''

    table3 = doc.add_table(rows=1, cols=5)
    table3.cell(0, 0).width = Cm(7)
    table3.cell(0, 0).text = 'Finance DCA'
    a = dictionary_1[name]['SRO Finance confidence']
    b = converting_RAGs(a)
    if a != None:
        table3.cell(0, 1).text = b
        cell_colouring(table3.cell(0, 1), b)
    else:
        table3.cell(0, 1).text = 'Not reported'

    try:
        a = dictionary_2[name]['SRO Finance confidence']
        b = converting_RAGs(a)
        if a != None:
            table3.cell(0, 2).text = b
            cell_colouring(table3.cell(0, 2), b)
        else:
            table3.cell(0, 2).text = 'Not reported'
    except KeyError:
        table3.cell(0, 2).text = 'N/A'

    try:
        a = dictionary_3[name]['SRO Finance confidence']
        b = converting_RAGs(a)
        if a != None:
            table3.cell(0, 3).text = b
            cell_colouring(table3.cell(0, 3), b)
        else:
            table3.cell(0, 3).text = 'Not reported'
    except KeyError:
        table3.cell(0, 3).text = 'N/A'

    try:
        a = dictionary_4[name]['SRO Finance confidence']
        b = converting_RAGs(a)
        if a != None:
            table3.cell(0, 4).text = b
            cell_colouring(table3.cell(0, 4), b)
        else:
            table3.cell(0, 4).text = 'Not reported'
    except KeyError:
        table3.cell(0, 4).text = 'N/A'

    '''SRO Benefits confidence'''

    table4 = doc.add_table(rows=1, cols=5)
    table4.cell(0, 0).width = Cm(7)
    table4.cell(0, 0).text = 'Benefits DCA'
    a = dictionary_1[name]['SRO Benefits RAG']
    b = converting_RAGs(a)
    if a != None:
        table4.cell(0, 1).text = b
        cell_colouring(table4.cell(0, 1), b)
    else:
        table4.cell(0, 1).text = 'Not reported'

    try:
        a = dictionary_2[name]['SRO Benefits RAG']
        b = converting_RAGs(a)
        if a != None:
            table4.cell(0, 2).text = b
            cell_colouring(table4.cell(0, 2), b)
        else:
            table4.cell(0, 2).text = 'Not reported'
    except KeyError:
        table4.cell(0, 2).text = 'N/A'

    try:
        a = dictionary_3[name]['SRO Benefits RAG']
        b = converting_RAGs(a)
        if a != None:
            table4.cell(0, 3).text = b
            cell_colouring(table4.cell(0, 3), b)
        else:
            table4.cell(0, 3).text = 'Not reported'
    except KeyError:
        table4.cell(0, 3).text = 'N/A'

    try:
        a = dictionary_4[name]['SRO Benefits RAG']
        b = converting_RAGs(a)
        if a != None:
            table4.cell(0, 4).text = b
            cell_colouring(table4.cell(0, 4), b)
        else:
            table4.cell(0, 4).text = 'Not reported'
    except KeyError:
        table4.cell(0, 4).text = 'N/A'

    '''SRO resourcing DCA'''

    table5 = doc.add_table(rows=1, cols=5)
    table5.cell(0, 0).width = Cm(7)
    table5.cell(0, 0).text = 'Resourcing DCA'
    a = dictionary_1[name]['Overall Resource DCA - Now']
    b = converting_RAGs(a)
    if a != None:
        table5.cell(0, 1).text = b
        cell_colouring(table5.cell(0, 1), b)
    else:
        table5.cell(0, 1).text = 'Not reported'

    try:
        a = dictionary_2[name]['Overall Resource DCA - Now']
        b = converting_RAGs(a)
        if a != None:
            table5.cell(0, 2).text = b
            cell_colouring(table5.cell(0, 2), b)
        else:
            table5.cell(0, 2).text = 'Not reported'
    except KeyError:
        table5.cell(0, 2).text = 'N/A'

    try:
        a = dictionary_3[name]['Overall Resource DCA - Now']
        b = converting_RAGs(a)
        if a != None:
            table5.cell(0, 3).text = b
            cell_colouring(table5.cell(0, 3), b)
        else:
            table5.cell(0, 3).text = 'Not reported'
    except KeyError:
        table5.cell(0, 3).text = 'N/A'

    try:
        a = dictionary_4[name]['Overall Resource DCA - Now']
        b = converting_RAGs(a)
        if a != None:
            table5.cell(0, 4).text = b
            cell_colouring(table5.cell(0, 4), b)
        else:
            table5.cell(0, 4).text = 'Not reported'
    except KeyError:
        table5.cell(0, 4).text = 'N/A'

    '''DCA Narrative text'''
    doc.add_paragraph()  # new
    y = doc.add_paragraph()
    heading = 'SRO DCA Narrative'
    y.add_run(str(heading)).bold = True

    dca_a = dictionary_1[name]['Departmental DCA Narrative']
    # print(dca_a)
    try:
        dca_b = dictionary_2[name]['Departmental DCA Narrative']
        # print(dca_b)
    except KeyError:
        dca_b = dca_a

    '''comparing text options'''
    # compare_text_showall(dca_a, dca_b, doc)
    compare_text_newandold(dca_a, dca_b, doc)

    '''Finance section'''
    y = doc.add_paragraph()
    heading = 'Financial information'
    y.add_run(str(heading)).bold = True

    '''Financial Meta data'''
    table1 = doc.add_table(rows=2, cols=5)
    table1.cell(0, 0).text = 'Forecast Whole Life Cost(£m):'
    table1.cell(0, 1).text = 'Percentage Spent:'
    table1.cell(0, 2).text = 'Source of Funding:'
    table1.cell(0, 3).text = 'Nominal or Real figures:'
    table1.cell(0, 4).text = 'Full profile reported:'

    table1.cell(1, 0).text = str(dictionary_1[name]['Total Forecast'])
    a = dictionary_1[name]['Total Forecast']
    b = dictionary_1[name]['Pre 18-19 RDEL Forecast Total']
    if b == None:
        b = 0
    # print(b)
    c = dictionary_1[name]['Pre 18-19 CDEL Forecast Total']
    if c == None:
        c = 0
    # print(c)
    d = dictionary_1[name]['Pre 18-19 Forecast Non-Gov']
    if d == None:
        d = 0
    # print(d)
    e = b + c + d
    try:
        c = round(e / a * 100, 1)
    except ZeroDivisionError:
        c = 0
    table1.cell(1, 1).text = str(c) + '%'
    a = str(dictionary_1[name]['Source of Finance'])
    b = dictionary_1[name]['Other Finance type Description']
    if b == None:
        table1.cell(1, 2).text = a
    else:
        table1.cell(1, 2).text = a + ' ' + str(b)
    table1.cell(1, 3).text = str(dictionary_1[name]['Real or Nominal - Actual/Forecast'])
    table1.cell(1, 4).text = 'complete manually'

    '''Finance DCA Narrative text'''
    doc.add_paragraph()
    y = doc.add_paragraph()
    heading = 'SRO Finance Narrative'
    y.add_run(str(heading)).bold = True

    '''RDEL'''
    rfin_dca_a = dictionary_1[name]['Project Costs Narrative RDEL']
    if rfin_dca_a == None:
        rfin_dca_a = 'None'

    try:
        rfin_dca_b = dictionary_2[name]['Project Costs Narrative RDEL']
        if rfin_dca_b == None:
            rfin_dca_b = 'None'
    except KeyError:
        rfin_dca_b = rfin_dca_a

    # compare_text_showall()
    compare_text_newandold(rfin_dca_a, rfin_dca_b, doc)

    '''CDEL'''
    cfin_dca_a = dictionary_1[name]['Project Costs Narrative CDEL']
    if cfin_dca_a == None:
        cfin_dca_a = 'None'

    try:
        cfin_dca_b = dictionary_2[name]['Project Costs Narrative CDEL']
        if cfin_dca_b == None:
            cfin_dca_b = 'None'
    except KeyError:
        cfin_dca_b = cfin_dca_a

    # compare_text_showall()
    compare_text_newandold(cfin_dca_a, cfin_dca_b, doc)

    '''financial chart heading'''  # new
    y = doc.add_paragraph()
    heading = 'Financial Analysis - Cost Profile'
    y.add_run(str(heading)).bold = True
    y = doc.add_paragraph()
    y.add_run('{insert chart}')

    '''milestone section'''
    y = doc.add_paragraph()
    heading = 'Planning information'
    y.add_run(str(heading)).bold = True

    '''Milestone Meta data'''
    table1 = doc.add_table(rows=2, cols=4)
    table1.cell(0, 0).text = 'Project Start Date:'
    table1.cell(0, 1).text = 'Latest Approved Business Case:'
    table1.cell(0, 2).text = 'Start of Operations:'
    table1.cell(0, 3).text = 'Project End Date:'

    c = dictionary_1[name]['Project MM18 Original Baseline']
    try:
        c = datetime.datetime.strptime(c.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
    except AttributeError:
        c = 'Not reported'
    table1.cell(1, 0).text = str(c)
    table1.cell(1, 1).text = str(dictionary_1[name]['BICC approval point'])
    a = dictionary_1[name]['Project MM20 Forecast - Actual']
    try:
        a = datetime.datetime.strptime(a.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
        table1.cell(1, 2).text = str(a)
    except AttributeError:
        table1.cell(1, 2).text = 'None'
    b = dictionary_1[name]['Project MM21 Forecast - Actual']
    try:
        b = datetime.datetime.strptime(b.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
    except AttributeError:
        b = 'Not reported'
    table1.cell(1, 3).text = str(b)

    # TODO: workout generally styling options for doc, paragraphs and tables
    # table1.style = "Heading1"

    '''milestone narrative text'''
    doc.add_paragraph()
    y = doc.add_paragraph()
    heading = 'SRO Milestone Narrative'
    y.add_run(str(heading)).bold = True

    mile_dca_a = dictionary_1[name]['Milestone Commentary']
    if mile_dca_a == None:
        mile_dca_a = 'None'

    try:
        mile_dca_b = dictionary_2[name]['Milestone Commentary']
        if mile_dca_b == None:
            mile_dca_b = 'None'
    except KeyError:
        mile_dca_b = mile_dca_a

    # compare_text_showall()
    compare_text_newandold(mile_dca_a, mile_dca_b, doc)

    '''milestone chart heading'''
    y = doc.add_paragraph()
    heading = 'Milestone Analysis - Swimlane Chart and Movements'
    y.add_run(str(heading)).bold = True
    y = doc.add_paragraph()
    y.add_run('{insert chart}')

    # doc.add_page_break()

    return doc


# doc = Document()

current_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_4_2018_wip.xlsx')
last_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_3_2018.xlsx')
Q2_ago_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_2_2018.xlsx')
Q3_ago_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_1_2017.xlsx')

current_Q_list = list(current_Q_dict.keys())
#current_Q_list = ['High Speed Rail Programme (HS2)']
# current_Q_list.remove('Commercial Vehicle Services (CVS)')

for x in current_Q_list:
    a = printing(x, current_Q_dict, last_Q_dict, Q2_ago_dict, Q3_ago_dict)
    a.save('C://Users//Standalone//Will//Q4_1819_{}_overview.docx'.format(x))

