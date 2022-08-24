import pandas

if __name__ == '__main__':
    htmlTables = pandas.read_html('ttt.htm')
    mWriter = pandas.ExcelWriter('out.xlsx')
    for i in range(len(htmlTables)):
        htmlTable = htmlTables[i]
        print(htmlTable.columns.values)
        # if len(htmlTable.columns) < 3 or htmlTable.columns.values
        #     continue
        print(htmlTable.columns)
        htmlTable.to_excel(mWriter, str(i))
    mWriter.close()
