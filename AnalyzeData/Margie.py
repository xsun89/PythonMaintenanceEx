import xlrd
import sets


# ----------------------------------------------------------------------
class ProcessIntercection:
    def __init__(self):
        self.path1 = ''
        self.path2 = ''
        self.keywordList1 = []
        self.keywordList2 = []
    def open_file(self, path):

        """
        Open and read an Excel file
        """
        try:
            book = xlrd.open_workbook(path)
            # get the first worksheet
            first_sheet = book.sheet_by_index(0)
            keywordList = []
            for r in range(1, first_sheet.nrows):
                keyword = str(first_sheet.cell(r, 2).value)
                keywords = keyword.split(';')
                for currentKeyword in keywords:
                    keywordForAdd = currentKeyword.strip().lower();
                    if keywordForAdd and keywordForAdd != 'keywords':
                        keywordList.append(keywordForAdd)

            return keywordList
        except:
            print 'An exception flew by!'
            raise


    def inputFileName(self, times):
        while True:
            try:
                path = raw_input("Please enter " + times + " Excel file name:")
                keywordList = self.open_file(path)
                if times == 'first':
                    self.path1 = path
                    self.keywordList1 = keywordList
                elif times == 'second':
                    self.path2 = path
                    self.keywordList2 = keywordList
                return
            except:
                print 'File name maybe wrong, please try again'


    def processData(self):
        try:
            set1 = sets.Set(self.keywordList1)
            set2 = sets.Set(self.keywordList2)
            keywordIntersection = set1.intersection(set2)
            print keywordIntersection
            filename = raw_input("Please enter file name to save the results:")
            f = open(filename, 'a')
            f.write(self.path1 + " and " + self.path2 + " keyword intersection start\n");
            f.writelines(["%s\n" % item for item in keywordIntersection])
            f.write(self.path1 + " and " + self.path2 + " keyword intersection end\n");
            return True
        except:
            print 'File name maybe wrong, please try again'
            return False
        finally:
            f.close()
            return True


# ----------------------------------------------------------------------
if __name__ == "__main__":
    ready = raw_input("Ready to start? Please type Yes or No or Exit")
    if ready.lower() != 'yes':
        exit()
    while True:
        myProcess = ProcessIntercection()
        myProcess.inputFileName('first')
        myProcess.inputFileName('second')
        myProcess.processData()
        next = raw_input("Continue to compare? Please type Yes or No")
        if next.lower() != 'yes':
            exit()

