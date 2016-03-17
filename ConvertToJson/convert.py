import xlrd
from collections import OrderedDict
import simplejson as json
 
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('MasterQuestionList.xls')
wb2 = xlrd.open_workbook('AnswerOptions.xls')
sh = wb.sheet_by_index(0)
sh2 = wb2.sheet_by_index(0)

jsondict = OrderedDict()
jsondict["Date"] = "12/8/2013"
jsondict["Version"] = 1

# List to hold dictionaries
questions_list = []
 
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    #for rownum2 in range (1, sh2.nrows):
    #Outter list (first file)
    questions = OrderedDict()
    row_values = sh.row_values(rownum)
    questions['QuestionNumber'] = int(row_values[0])
    questions['QuestionVersion'] = row_values[1]
    questions['QuestionCategory'] = row_values[2]
    questions['QuestionTitle'] = row_values[3]
    questions['QuestionText'] = row_values[4]
    questions['QuestionType'] = row_values[5]
    
    for rownum2 in range(1, sh2.nrows):
    #Inner list (second file)
        row_values2 = sh2.row_values(rownum2)
        if (int(row_values2[0]) != questions['QuestionNumber']):
            continue
        innerMap = OrderedDict()
        innerMap["QuestionNumber"] = int(row_values2[0])
        innerMap["QuestionVersion"] = row_values2[1]
        innerMap["OptionTitle"] = row_values2[2]
        innerMap["OptionText"] = row_values2[3]
        innerMap["Points"] = row_values2[4]

        if 'AnswerOptions' not in questions:
            questions['AnswerOptions'] = []
        questions['AnswerOptions'].append(innerMap)
    
    
#    questions_list.append(innerMap)
    questions_list.append(questions)
 
# Serialize the list of dicts to JSON

jsondict["Questions"] = questions_list
j = json.dumps(jsondict)
 
# Write to file
with open('questions.json', 'w') as f:
    f.write(j)
