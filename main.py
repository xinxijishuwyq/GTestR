import collections

from docxtpl import DocxTemplate
import pandas


class MFiles(object):
    def __init__(self, m_nameLine, num):
        self.nameLine = m_nameLine
        t = self.nameLine.split(':')
        self.name = t[0]
        self.line = t[1]
        self.num = num
        self.nNameLine = num + '. ' + self.nameLine


class MTable(object):

    def __init__(self):
        self.strNames = ''
        self.mFiles = []
        self.count = 0
        self.name = set()
        self.chs = ''

    def add(self, m_nameLine, m_chs):
        self.count += 1
        self.mFiles.append(MFiles(m_nameLine, str(self.count)))
        self.name.add(self.mFiles[self.count - 1].name)
        self.strNames = '、'.join(self.name)
        self.chs = m_chs


if __name__ == '__main__':
    mYdict = collections.defaultdict(lambda: MTable())
    mChsdict = collections.defaultdict(str)
    # mYdict['GJB1111'].add('mdzz.cpp: 174')
    # mYdict['GJB1111'].add('mdzz.cpp: 179')
    # mYdict['GJB2113'].add('naizi.cpp: 233')
    # mYdict['GJB2113'].add('naddizi.cpp: 233')
    df = pandas.read_excel('test.xlsx', sheet_name=0)
    codeList = df['code'].tolist()
    chsList = df['chinese'].tolist()
    nameLineList = df['nameLine'].tolist()
    coreExcelData = list(zip(codeList, chsList, nameLineList))
    for i in coreExcelData:
        mChsdict[i[0]] = i[1]
        mYdict[i[0]].add(i[2], i[1])
    tpl = DocxTemplate('template.docx')
    context = {
        'myTables': mYdict
    }
    tpl.render(context)
    tpl.save('out.docx')
