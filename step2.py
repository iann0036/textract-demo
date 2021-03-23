import io
import json
import xlwt

OUTPUT_XLS = "output.xls"

analysis_result = {}
with open("analysisresult.json", "r") as f:
    analysis_result = json.loads(f.read())

##### STEP 2. Extract table information into Excel workbook

book = xlwt.Workbook()

table_number = 1

for block in analysis_result['Blocks']:
    if block['BlockType'] == 'TABLE': # document tables

        sheet = book.add_sheet("Table {}".format(table_number))
        table_number+=1
        table = {}

        for relationship in block['Relationships']:
            if relationship['Type'] == 'CHILD':
                for rid in relationship['Ids']:
                    for block2 in analysis_result['Blocks']: # table cells
                        if block2['Id'] == rid:
                            row = block2['RowIndex'] - 1
                            column = block2['ColumnIndex'] - 1
                            if 'Relationships' in block2:
                                for relationship2 in block2['Relationships']:
                                    if relationship2['Type'] == 'CHILD':
                                        for rid2 in relationship2['Ids']:
                                            for block3 in analysis_result['Blocks']: # cell words
                                                if block3['Id'] == rid2:
                                                    if block3['BlockType'] == 'WORD':
                                                        word = block3['Text']
                                                        
                                                        if row not in table:
                                                            table[row] = {}
                                                        
                                                        if column in table[row]:
                                                            table[row][column] += " " + word
                                                        else:
                                                            table[row][column] = word

        for rowindex, row in table.items():
            for columnindex, text in row.items():
                sheet.write(rowindex, columnindex, text)

book.save(OUTPUT_XLS)
