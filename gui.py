from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import re

from form2doc import form2doc, get_forms

# Dictionary mapping input form paths to PyPDF form-field dictionaries
formpaths_to_formdicts = {}

template_form_path = None
# PdfFileReader object
template_form = None
# Forms dictionary
template_fields = None
# Dictionary of Output Templates to lists of rules, each with at least one string pointers
# to potentit\al worker form field, and one to an output template field
out_tmplts_to_rules = {}


def add_rule():
    pass

def save_formedit():
    f = app.getTabbedFrameSelectedTab('Worker Forms')
    print(f)
    worker_entries = formpaths_to_formdicts[f]
    fields = {}
    # fields0=fields
    print('hey')
    for k in worker_entries.keys():
        print(k + ' : '+app.getEntry(k+f))
        fields[k]= app.getEntry(k+f)
        # fields0[k]=''
    outform = PdfFileWriter()
    original_form = PdfFileReader(file(f,'rb'))
    outStream = file('testout.pdf', 'wb')
    # outform.write(outStream)
    # OUTTEST = file('supertesty.pdf','wb')
    # outform.write(OUTTEST)
    # OUTTEST.close()
    print(outform.getNumPages())
    print(range(outform.getNumPages()))

    for p in range(original_form.getNumPages()):
        outform.addPage(original_form.getPage(p))
    for p in range(outform.getNumPages()):
        print(p)
        print(fields)
        # outform.updatePageFormFieldValues(outform.getPage(p), fields0)
        outform.updatePageFormFieldValues(outform.getPage(p), fields)
    outform.write(outStream)
    outform.write(outStream)
    # outform.write(outStream)
    # outform.write(outStream)

    outStream.close()

def add_rule( ):
    pass

def get_files(filedir):
    i=0
    for f in filedir:
        if os.path.isfile(f) and (f.endswith('.pdf')):
            print(f +' is a file')
        elif os.path.isdir(f):
            out = []
            for j in os.listdir(f):
                out = out + get_files([f+'/'+j])
            filedir= filedir[:i]+out+filedir[i+1:]
        else: filedir=filedir[:i]+filedir[i+1:]
        i+=1
    return filedir

def addIn(data):
    global formpaths_to_formdicts

    # drop input processing into list of files
    data = data.split(' ')
    data0 = data
    i = 0
    for d in data:
        if not (d.startswith('/') or d.startswith('{')):
            data0[i-1] = data0[i-1]+' '+d
        else: data0=data0+[d]
        i+=1
    app.setStretch('column')
    app.openTabbedFrame('Worker Forms')
    data =[d.strip('{').strip('}') for d in data0]
    files = []
    for f in data:
        files = files + get_files([f])

    for f0 in files:
        print(f0)
        f=f0.decode('utf-8')
        print(f)
        if f not in app.getAllListItems('inputs'):
            app.addListItem('inputs', f)
            # worker_entries = get_forms(f, inputs_to_form_entries)
            worker_entries= PdfFileReader(file(f,'rb')).getFields()
            formpaths_to_formdicts[f] = worker_entries
            try:
                worker_fullname = worker_entries['Last']['/V']
            except KeyError:
                worker_fullname = f.split('/')[-1]
            except TypeError:
                worker_fullname = f.split('/')[-1]
            app.setStretch('both')
            with app.tab(f0):
                app.setTabText('Worker Forms', f0, worker_fullname)

                with app.scrollPane('Worker Form Edit' + f):
                    i=0
                    for key in worker_entries.keys():
                        # default=False
                        # if '(' in key:
                        #     desired_format = key.split('(')[1].strip(')')
                        #     key_0=key.split('(')[0]
                        #     default=True
                        # else: key_0=key
                        app.setSticky('e')

                        app.addLabel(key+f,key[:25],i,2,0)
                        app.setLabelTooltip(key+f,key)
                        # app.setStretch('both')
                        app.setSticky('w')
                        app.addEntry(key+f,i,3,1,0)
                        # if default:
                        #     app.setEntryDefault(key_0+f,desired_format)

                        try:
                            app.setEntry(key+f,worker_entries[key]['/V'])
                        except TypeError:
                            print('no '+key+ ' form in '+ f)
                        # app.addScrolledTextArea(key+f,i,4)

                        i+=1



def remove_input():
    selected_form = app.getListBox('inputs')
    app.removeListItem('inputs',selected_form)
    app.deleteTabbedFrameTab('Worker Forms', selected_form[0])
    for w in app.widgetManager.group(app.Widgets.Label):
        if selected_form[0] in w:
            print(app.widgetManager.get(app.Widgets.Label,w))
            app.widgetManager.get(app.Widgets.Label,w).destroy()
    for w in app.widgetManager.group(app.Widgets.Entry):
        if selected_form[0] in w:
            app.widgetManager.get(app.Widgets.Entry,w).destroy()


def save_form_template():
    global template_form
    global template_fields
    global template_form_path
    writer_copy=PdfFileWriter()
    writer_copy.cloneReaderDocumentRoot(template_form)
    for p in range(writer_copy.getNumPages()):
        page = writer_copy.getPage(p)
        print(page)
        writer_copy.updatePageFormFieldValues(page, template_fields)

    writer_copy.write(open(template_form_path.strip('.pdf') + '_copy_TEST.pdf','wb'))

def set_form_template():
    global template_fields
    global template_form
    global template_form_path

    app.setStretch('column')


    template_form_path = app.openBox('form_template')
    app.setLabel('form_template_path', template_form_path)
    template_form = PdfFileReader(template_form_path)

    template_fields = template_form.getFields()
    form_entries_text = ''
    for f in template_fields:
        form_entries_text = form_entries_text+'\n'+f
        app.addListItem('form_template_list',f)




def update_tab_select():
    if app.getListBox('inputs')!= None:
        selected_form = app.getListBox('inputs')[0]
        selected_tab = app.getTabbedFrameSelectedTab('Worker Forms')
        if selected_tab != selected_form:
            app.setTabbedFrameSelectedTab('Worker Forms', selected_form)

def update_form_path_select():
    selected_form = app.getListBox('inputs')[0]
    selected_tab = app.getTabbedFrameSelectedTab('Worker Forms')
    if not selected_tab == selected_form:
        app.selectListItem('inputs',selected_tab)

selected_form_template_entry = None


def update_form_template_edit():
    app.setEntry('form_entry_edit',app.getListBox('form_template_list')[0])
    global selected_form_template_entry
    selected_form_template_entry = app.getListBoxPos('form_template_list')[0]


def update_form_template_list():
    global selected_form_template_entry
    global template_fields

    entry = app.getEntry('form_entry_edit')
    app.selectListItemAtPos('form_template_list',selected_form_template_entry)
    old_entry = app.getListBox('form_template_list')[0]
    if  old_entry != entry:
        # nested dictionary with title, value info
        formbox = template_fields.pop(old_entry)
        print(formbox)
        print(entry)
        template_fields[entry] = formbox
        app.setListItemAtPos('form_template_list',selected_form_template_entry,entry)

    # print(template_fields)
def save_form():
    pass


with gui("OPC form2doc") as app:
    app.setStretch('both')
    app.addListBox('inputs',[],1,0,1,30)
    app.setListBoxChangeFunction('inputs', update_tab_select)
    with app.tabbedFrame('Worker Forms',0,1,2,31):
        with app.scrollPane('worker_form_scroll',0,1,1,31):
            app.setListBoxDropTarget('inputs', addIn, replace=False)
            app.setTabbedFrameChangeFunction('Worker Forms',update_form_path_select)

    app.setStretch('column')
    app.addLabel('Drop Inputs', 'Drop Inputs',0,0,1,1)
    app.addButton('Remove Input',remove_input,32,0)
    app.addButton('Save Changes to Form',save_formedit,32,1)

    app.addLabel('Select Outputs','Select Outputs',0,2)
    app.addListBox('outputs',[],1,2,1,31)
    # with app.scrollPane('form_template_scroll',3,4):

    # app.addButton('Open', set_form_template,1,2)
    # app.addListBox('form_template_list',[],3,2,1,1)


    # app.addLabel('Form Template','Form Template',0,2)
    # app.addLabel('form_template_path','',2,2)
    # app.addEntry('form_entry_edit',4,2)
    # app.setListBoxChangeFunction('form_template_list',update_form_template_edit)
    # app.setEntrySubmitFunction('form_entry_edit',update_form_template_list)
    # app.addButton('Save', save_form_template,5,2)

app.setFont(15)
app.setBg("black")
app.setFg("lightGray")



app.go()
