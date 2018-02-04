import sys
from PyPDF2 import PdfFileWriter, PdfFileReader
from docx import Document
import re
import sys
reload(sys)
sys.setdefaultencoding('utf8')

# Loop through forms supplied as arguments
for i in sys.argv[1:]:
    # extract information from PDF form
    form = PdfFileReader(i).getFields()
    # keys = ['JOB NAMECOMPANY','FULL NAME AS APPEARS ON PASSPORT','SEX','DATE OF BIRTH MMDDYYYY','SOCIAL SECURITY NUMBER','ADDRESS','TELEPHONE NUMBER','POSITION  JOB TITLE','EXPECTED TOTAL WAGES FOR THIS JOB','COUNTRY OF BIRTH']
    # info = [form[k] for k in keys]
    fullname = form['FULL NAME AS APPEARS ON PASSPORT']['/V']
    position = form['POSITION  JOB TITLE']['/V']
    jobname_company = form['JOB NAMECOMPANY']['/V']
    jobname_company = re.split('(/)', jobname_company)
    # first element is the job name, second is the company
    jobname=jobname_company[0]
    company= jobname_company[-1]
    sex = form['SEX']['/V']
    dob = form['DATE OF BIRTH MMDDYYYY']['/V']
    sin = form['SOCIAL SECURITY NUMBER']['/V']
    addy= form['ADDRESS']['/V']
    tel = form['TELEPHONE NUMBER']['/V']
    shootdates = form['DATE OF JOB SHOOTING']['/V']
    shootdates = re.split('(to|-)',shootdates)
    startdate = shootdates[0]
    enddate = shootdates[-1]
    birthdate = form['DATE OF BIRTH MMDDYYYY']['/V']
    immigration_processing = form['IMMIGRATION PROCESSING AT']['/V']
    birthcountry = form['COUNTRY OF BIRTH']['/V']

    # The replacement dictionary
    unpaid_replace = {'[name, title]': fullname +', '+position,
    '[name]': fullname,
    '[name here]': fullname,
    '[client]': fullname,
    '[visitor name]' : fullname,
    '[visitor name' : fullname,

    'dob':'',
    'm/d/y': birthdate,

    '[position]': position,
    '[foreign production company]':company,
    '[producer]': company,
    '[city]':immigration_processing,
    '[name of job]': jobname,
    '[production]': jobname,
    '[job name]': jobname,
    '[job]': jobname,
    '[start date]':startdate,
    '[end date]': enddate,
    '[budget]': 'Budget?',
    '[nationality]': birthcountry,
    '[position]': position,
    '[foreign production company]': 'Foreign?',
    '[days in canada]': 'DayCalculation',

    '[agency or client]': 'Agency?',
    '[agency]':'Agency?',



    '[shoot dates]':startdate+' and '+enddate,


    u'[description of client\u2019s business]':'a smuggler',

    '[location]':'Location?',
    '[production dates]': startdate+' and '+enddate,
    ']':''
    }

    # write UNPAID .docx file with client name
    invite = Document('InvitationLetterTemplate.docx')
    for p in invite.paragraphs:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                print(inline[i].text)
                m = re.findall('(\[.*?\]|\[.*|\]|DOB M\/D\/Y)', inline[i].text)
                for replacee in m:
                    print(replacee)
                    replacee_0 = replacee.lower()
                    new_text = inline[i].text.replace(replacee, unpaid_replace[replacee_0])
                    inline[i].text = new_text
    invite.save(fullname +' unpaid.docx')

    # Write PAID .docx file w/ client info
    invite2 = Document('/Users/sme/Desktop/form2doc/OPCBVLetter_Blank.docx')
    for p in invite2.paragraphs:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                # print(inline[i].text)
                m = re.findall('(\[.*?\]|\[.*|\]|DOB M\/D\/Y)', inline[i].text)
                for replacee in m:
                    # print(replacee)
                    replacee_0 = replacee.lower()
                    new_text = inline[i].text.replace(replacee, unpaid_replace[replacee_0])
                    inline[i].text = new_text
    invite2.save(fullname +' paid.docx')

### FOR DEBUGGING

# # write .docx file with client name
# unpaid = Document('/Users/sme/Desktop/form2doc/OPCBVLetter_Blank.docx')
# for p in unpaid.paragraphs:
#     inline = p.runs
#     # Loop added to work with runs (strings with same style)
#     for i in range(len(inline)):
#         m = re.findall('(\[.*?\])', inline[i].text)
#         for mi in m:
#             print(mi[1:-1])
