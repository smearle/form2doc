from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader
from docx import Document
import os
import re

from form2doc import form2doc

# A dictionary pointing distinct keys for each form (based and progressively more
# fine-grained information) to their filepaths
inkeys2paths = {}

src_dir = os.path.dirname(os.path.realpath(__file__))
# OPC entry letter rules
rulesheet_dirpath = src_dir + '/Rules'
input_templates_dirpath = src_dir + '/Input Templates'
output_templates_dirpath = src_dir + '/Output Templates'

input_dirpath = src_dir + '/Input'
output_dirpath = src_dir + '/Output'

# Dictionary mapping input form paths to PyPDF form-field dictionaries
formpaths_to_formdicts = {}
# Doctionary mapping input form paths to (title to entry) form-field dictionaries
formpaths_to_fielddicts ={}

template_form_path = None
# PdfFileReader object
template_form = None
# Forms dictionary
template_fields = None
# Dictionary of out_templates to lists of rules, each with at least one string pointers
# to potentit\al worker form field, and one to an output template field
# out_template --> [fillable_field(s)] --> [replacement(s)] --> [condition(s)] (optional)
#                                      --> alt [replacement(s)] --> alt [conditions]
outfields2rules = {}
# Dictionary of input paths to lists of output paths
ins_to_outs = {}

box_ptrn = re.compile('\[.*?\]')

def addRule():
    pass

def saveFormedit():
    f = app.getTabbedFrameSelectedTab('inputs')
    worker_entries = formpaths_to_formdicts[f]
    fields = formpaths_to_fielddicts[f]
    # fields0=fields
    for k in worker_entries.keys():
        fields[k]= app.getEntry(k+f)
        # fields0[k]=''
    outform = PdfFileWriter()
    original_form = PdfFileReader(file(f,'rb'))
    outStream = file('testout.pdf', 'wb')
    # outform.write(outStream)
    # OUTTEST = file('supertesty.pdf','wb')
    # outform.write(OUTTEST)
    # OUTTEST.close()

    for p in range(original_form.getNumPages()):
        outform.addPage(original_form.getPage(p))
    for p in range(outform.getNumPages()):
        # outform.updatePageFormFieldValues(outform.getPage(p), fields0)
        outform.updatePageFormFieldValues(outform.getPage(p), fields)
    outform.write(outStream)
    outform.write(outStream)
    # outform.write(outStream)
    # outform.write(outStream)

    outStream.close()

def addRule( ):
    pass

def getFiles(filedir, extension):
    i=0
    for f in filedir:
        if os.path.isfile(f) and (f.endswith('.' + extension)):
            pass
        elif os.path.isdir(f):
            out = []
            for j in os.listdir(f):
                out = out + getFiles([f+'/'+j],extension)
            filedir= filedir[:i]+out+filedir[i+1:]
        else: filedir=filedir[:i]+filedir[i+1:]
        i+=1
    return filedir

def parseDropDate(data):
    # drop input processing into list of files
    print(data)
    data = data.split(' ')
    data0 = data
    i = 0
    for d in data:
        if not (d.startswith('/') or d.startswith('{')):
            data0[i-1] = data0[i-1]+' '+d
        i+=1
    data =[d.strip('{').strip('}') for d in data0]
    print(data)

    return data

def addInDrop(dropdata):
    global formpaths_to_formdicts
    data =parseDropDate(dropdata)
    app.setStretch('column')
    app.openTabbedFrame('inputs')
    files = []
    for f in data:
        files = files + getFiles([f],'pdf')
    for f in files:
        f=f.decode('utf-8')
        addIn(f)

def addIn(filepath):
    app.openTabbedFrame('inputs')
    if filepath not in app.getAllListItems('inputs'):
        app.addListItem('inputs', filepath)
        # worker_entries = get_forms(f, inputs_to_form_entries)
        worker_entries= PdfFileReader(file(filepath,'rb')).getFields()
        formpaths_to_formdicts[filepath] = worker_entries
        fields = {}
        for k in worker_entries.keys():
            fields[k]= worker_entries[k]['/V']
        formpaths_to_fielddicts[filepath] = fields
        try:
            fullname = worker_entries['Last']['/V']
        except KeyError:
            fullname = filepath.split('/')[-1]
        except TypeError:
            fullname = filepath.split('/')[-1]

        # # TODO Globalize
        # runnerup_fields = ['POSITION  JOB TITLE', 'Email']
        # for r in runnerup_fields:
        #     try:
        #         if (key in inkeys2paths) and (worker_entries[r]['/V'] != None) and (worker_entries[r]['/V'] != ''):
        #             key += '_' + worker_entries[r]['/V']
        #     except KeyError:
        #         print(r +' field not found in '+ filepath)
        # # TODO still not unique..?
        # inkeys2paths[key] = filepath


        app.setStretch('both')
        with app.tab(filepath):
            app.setTabText('inputs',filepath,fullname)
            with app.scrollPane('Worker Form Edit' + filepath):
                i=0
                for key in worker_entries.keys():
                    app.setSticky('e')

                    app.addLabel(key+filepath,key[:25],i,2,0)
                    app.setLabelTooltip(key+filepath,key)
                    # app.setStretch('both')
                    app.setSticky('w')
                    app.addEntry(key+filepath,i,3,1,0)
                    # if default:
                    #     app.setEntryDefault(key_0+f,desired_format)

                    try:
                        app.setEntry(key+filepath,worker_entries[key]['/V'])
                    except TypeError:
                        pass
                    except KeyError:
                        pass
                    # app.addScrolledTextArea(key+f,i,4)

                    i+=1

def addOut(filepath):
    pass

def removeInput():
    selected_form = app.getListBox('inputs')
    app.removeListItem('inputs',selected_form)
    app.deleteTabbedFrameTab('inputs', selected_form[0])
    for w in app.widgetManager.group(app.Widgets.Label):
        if selected_form[0] in w:
            app.widgetManager.get(app.Widgets.Label,w).destroy()
    for w in app.widgetManager.group(app.Widgets.Entry):
        if selected_form[0] in w:
            app.widgetManager.get(app.Widgets.Entry,w).destroy()

def removeOutputTemplate():
    app.deleteTabbedFrameTab('out_templates',app.getTabbedFrameSelectedTab('out_templates'))


def saveFormTemplate():
    global template_form
    global template_fields
    global template_form_path
    writer_copy=PdfFileWriter()
    writer_copy.cloneReaderDocumentRoot(template_form)
    for p in range(writer_copy.getNumPages()):
        page = writer_copy.getPage(p)
        writer_copy.updatePageFormFieldValues(page, template_fields)

    writer_copy.write(open(template_form_path.strip('.pdf') + '_copy_TEST.pdf','wb'))


def updateInputTabSelect():
    if app.getListBox('inputs')!= None:
        selected_form = app.getListBox('inputs')[0]
        selected_tab = app.getTabbedFrameSelectedTab('inputs')
        if selected_tab != selected_form:
            app.setTabbedFrameSelectedTab('inputs', selected_form)



def updateFormPathSelect():
    selected_form = app.getListBox('inputs')[0]
    selected_tab = app.getTabbedFrameSelectedTab('inputs')
    if not selected_tab == selected_form:
        app.selectListItem('inputs',selected_tab)

selected_form_template_entry = None


def updateFormTemplateEdit():
    app.setEntry('form_entry_edit',app.getListBox(app.getTabbedFrameSelectedTab('form_templates'))[0])
    # global selected_form_template_entry
    # selected_form_template_entry = app.getListBoxPos('form_templates')[0]


def updateFormTemplates():
    global selected_form_template_entry
    global template_fields

    entry = app.getEntry('form_entry_edit')
    app.selectListItemAtPos('form_templates',selected_form_template_entry)
    old_entry = app.getListBox('form_templates')[0]
    if  old_entry != entry:
        # nested dictionary with title, value info
        formbox = template_fields.pop(old_entry)
        template_fields[entry] = formbox
        app.setListItemAtPos('form_templates',selected_form_template_entry,entry)

def saveForm():
    pass

def generateOutput():
    sel_inpath = app.getListBox('inputs')
    out_template = app.openBox()

def updateRuleeditEntry():
    app.setEntry('rule_edit',app.getListBox(app.getTabbedFrameSelectedTab('rulesheets')+'_rules')[0])

def addInTemplatesDrop(dropdata):
    data = parseDropDate(dropdata)
    app.openTabbedFrame('form_templates')
    files = []
    for f in data:
        files = files + getFiles([f],'pdf')
    for f in files:
        addInTemplate(f)

def addInTemplate(filepath):
    app.openTabbedFrame('form_templates')
    template_name = filepath.split('/')[-1].replace('.pdf','')
    template_form = PdfFileReader(file(filepath,'rb'))
    with app.tab(template_name):
        app.setStretch('both')

        app.addListBox(template_name,template_form.getFields(),0,0,10,10)
        app.setListBoxGroup(template_name)
        app.setListBoxChangeFunction(template_name,updateFormTemplateEdit)



def addOutTemplatesDrop(dropdata):
    data = parseDropDate(dropdata)
    files = []
    for f in data:
        files = files + getFiles([f],'docx')
    for f in files:
        add_out_template(f)


def addOutTemplate(f):
    app.setStretch('both')
    app.openTabbedFrame('out_templates')
    # Assumes the name (not just path) of each template file is unique.
    template_name = f.split('/')[-1].replace('.docx','')

    with app.tab(f):
        app.setTabText('out_templates',f,template_name)
        # app.addListBox(template_name + '_boxes',boxes,0,2)
        app.addListBox(f ,None,0,1,10,10)
        app.setListBoxGroup(f)
        app.setListBoxChangeFunction(f,updateRuleeditEntry)
        outfields = getDocFields(f)
        for outfield in outfields:
            if outfield not in app.getAllListItems(f):
                app.addListItem(f, outfield)

        # add empty rules for unaccounted-for outfields
        rule_outfields = []
        sel_rulesheet = app.getTabbedFrameSelectedTab('rulesheets')
        for rule in app.getAllListItems(sel_rulesheet+'_rules'):
            rule_outfields += box_ptrn.findall(rule.split('] replaced by [')[0])
        for outfield in app.getAllListItems(f):
            if outfield not in rule_outfields:
                app.addListItem(sel_rulesheet+'_rules', outfield +' replaced by []')


def getDocFields(docpath):
    # Returns a list of fields to be replaced in a text-based document, where
    # such fields are denoted by brackets "[]" surrounding text describing
    # info to be filled in.
    out_doc = Document(docpath)
    fillable_fields = []
    j=0
    for p in out_doc.paragraphs:
        inline = p.runs
        par_text =''
        # Assembles paragraph-text across differently-styled runs.
        for i in range(len(inline)):
            par_text = par_text + inline[i].text
        boxes = re.findall('(\[.*?\])', par_text)
        for box in boxes:
            fillable_fields = fillable_fields + [box.encode('utf8')]
        j=j+1
    return fillable_fields

def updateRules():
    sel_rulesheet = app.getTabbedFrameSelectedTab('rulesheets')
    sel_rule = app.getListBox(sel_rulesheet+'_rules')
    new_rule = app.getEntry('rule_edit')
    app.setListItemAtPos(sel_rulesheet+'_rules',app.getListBoxPos(sel_rulesheet+'_rules')[0],new_rule)
    app.selectListItem(sel_rulesheet+'_rules', app.getEntry('rule_edit'))

def addCondition():
    rule = app.getEntry('rule_edit')
    if '], when ' in rule:
        subrule = re.search('\], when (\[.*?\] .+? \[.*?\], )+or \[.*?\] .+? \[.*?\]|\], when (\[.*?\] .+? \[.*?\])',rule).group(0)
        new_sub = subrule.replace('], or [','], [') + ', or [] == []'
        new_rule = rule.replace(subrule, new_sub)
    else:
        new_rule = rule + ', when [] == []'
    app.setEntry('rule_edit', new_rule)

def addReplacee():
    rule = app.getEntry(rule_edit)
    p = rule.find(' replaced by')
    app.setEntry(rule_edit, rule[:p]+', []')

def addReplacement():
    rule = app.getEntry('rule_edit')
    subrule = re.search('(\] replaced by \[.*?\], )+or \[.*?\]|\] replaced by (\[.*?\])',rule).group(0)
    p = rule.find(subrule)
    new_sub = subrule.replace('], or [','], [') + ', or []'
    new_rule = rule.replace(subrule, new_sub)
    app.setEntry('rule_edit', new_rule)
    updateRules()

def pasteEntry2Rule():
    src_entry, trg_entry = 'form_entry_edit','rule_edit'
    app.setEntry(trg_entry,
    # Replaces the first empty box in the rule with the selected form template field.
    re.split('(\[\])',app.getEntry(trg_entry),1)[0] + '['+app.getEntry(src_entry)
                            +']'+re.split('\[\]',app.getEntry(trg_entry),1)[1])
    updateRules()

def delEntryFromRule():
    trg = 'rule_edit'
    app.setEntry(trg, re.sub('\][^\[]+?\[ ', '][ ',app.getEntry(trg)[::-1],1)[::-1])
    updateRules()

def saveRulesheet():
    rulesheet = app.getTabbedFrameSelectedTab('rulesheets')
    with open(dir_path+'/Rules/'+rulesheet+".txt", "w") as text_file:
        txt= ''
        for t in app.getAllListItems(rulesheet+'_rules'):
            txt += '\n'+t
        text_file.write(txt)

def paste_outfield():
    rule = app.getEntry('rule_edit')
    subrule  = rule.split('] replaced by [')[0] +']'
    outfield = app.getListBox(app.getTabbedFrameSelectedTab('out_templates'))[0]
    if ('], or [' in subrule) or ('] or [' in subrule):
        new_sub = subrule.replace('], or [', '], [').replace('] or [', '], [')
        new_sub = new_sub +', or ' + outfield
    else:
        new_sub = subrule + ' or ' + outfield
    new_rule = rule.replace(subrule, new_sub)
    app.setEntry('rule_edit', new_rule)

def addRulesheetsDrop(dropdata):
    data = parseDropDate(dropdata)
    files = []
    for f in data:
        files += getFiles(f)
    for f in files:
        addRulesheet(f)

def addRulesheet(path):
    rules = open(path,'rb').readlines()
    sheetname = path.split('/')[-1].replace('.txt','')
    app.openTabbedFrame('rulesheets')
    with app.tab(sheetname):
        app.addListBox(sheetname+'_rules', rules, 0,0)
        app.setListBoxGroup(sheetname+'_rules')
        app.setListBoxChangeFunction(sheetname+'_rules',updateRuleeditEntry)

with gui("OPC form2doc") as app:
    app.setStretch('both')
    app.addListBox('inputs',[],1,0,1,31)
    app.setListBoxChangeFunction('inputs', updateInputTabSelect)
    with app.tabbedFrame('inputs',1,1,1,31):
        app.setListBoxDropTarget('inputs', addInDrop, replace=False)
        app.setTabbedFrameChangeFunction('inputs',updateFormPathSelect)

    app.addListBox('outputs',[],1,2,1,31)
    with app.tabbedFrame('output_preview',1,3):
        pass
    app.setStretch('column')
    app.addLabel('Input Preview', 'Input Preview', 0,1)
    app.addLabel('Inputs', 'Inputs',0,0,1,1)
    app.addButton('Remove Input',removeInput,32,0)
    app.setButtonTooltip('Remove Input', 'Does not delete input form from hard drive.')
    app.addButton('Save Form',saveFormedit,32,1)

    app.addLabel('Outputs','Outputs',0,2)
    app.addButton('Generate Output',generateOutput,32,2)

    app.addHorizontalSeparator(33,0,4)
    app.addLabel('out_templates','Output Templates',34,3)
    app.addLabel('Output Preview','Output Preview',0,3)
    app.addButton('Save Output',None,32,3)
    app.setStretch('both')
    with app.tabbedFrame('rulesheets',35,1,2,30):
        app.setTabbedFrameDropTarget('rulesheets', addRulesheetsDrop)


    with app.tabbedFrame('out_templates',35,3):
        app.setTabbedFrameDropTarget('out_templates', addOutTemplatesDrop)
    app.setStretch('column')
    app.addEntry('rule_edit',66,1,2,1)
    app.setEntrySubmitFunction('rule_edit',updateRules)
    app.addButton('delete last',delEntryFromRule,67,1)
    app.addButton('Add Condition',addCondition,68,2)
    app.addButton('Add Replacement',addReplacement,68,1)
    app.addButton('Save Rule Sheet',saveRulesheet,69,2)

    app.addButton('Remove Output Template',removeOutputTemplate,69,3)
    app.setButtonTooltip('Remove Output Template','Does not delete output template from hard drive.')
    app.addLabel('Form Template','Input Templates',34,0)
    app.addLabel('Rule Sheets','Rule Sheets',34,1,2,1)
    app.addEntry('form_entry_edit',66,0)
    app.addButton('>',pasteEntry2Rule,67,0)
    app.addButton('Save to Form Template', saveFormTemplate,68,0)
    app.addButton('<',paste_outfield,66,3)
    app.setStretch('both')
    with app.tabbedFrame('form_templates',35,0,1,30):
        app.setTabbedFrameDropTarget('form_templates', addInTemplatesDrop)

    # Loading default templates and rules
    for path in os.listdir(rulesheet_dirpath):
        if path.endswith('.txt'):
            addRulesheet(rulesheet_dirpath+'/'+path)
    for path in os.listdir(input_templates_dirpath):
        if path.endswith('.pdf'):
            addInTemplate(input_templates_dirpath+'/'+path)
    for path in os.listdir(output_templates_dirpath):
        if path.endswith('.docx'):
            addOutTemplate(output_templates_dirpath+'/'+path)
    # Load default input folder
    for path in os.listdir(input_dirpath):
        if path.endswith('.pdf'):
            addIn(input_dirpath+'/'+path)


    # app.setEntrySubmitFunction('form_entry_edit',updateFormTemplates)




app.setFont(15)
app.setBg("black")
app.setFg("lightGray")



app.go()
