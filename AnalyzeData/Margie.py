import xlrd
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    # print number of sheets
    print book.nsheets
    # print sheet names
    sheet = book.sheet_names()
    print type(sheet)
    print book.sheet_names()
    print sheet[0]
    sheet_n = sheet[0]
    print type(sheet[0])
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    # read a row
    print first_sheet.row_values(0)
    # read a cell
    cell = first_sheet.cell(0,0)
    print cell
    print cell.value
    # read a row slice
    print first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2)
    #sheet = book.sheet_by_name("source")
    for r in range(1, sheet_n.nrows):
        product      = sheet_n.cell(r,).value
        print product
        customer = sheet_n.cell(r,1).value
        print customer
        rep = sheet_n.cell(r,2).value
        print rep
# Establish a MySQL connection
#database = MySQLdb.connect (host="localhost", user = "root", passwd = "", db = "mysqlPython")

# Get the cursor, which is used to traverse the database, line by line
#cursor = database.cursor()

# Create the INSERT INTO sql query
#query = """INSERT INTO orders (product, customer_type, rep, date, actual, expected, open_opportunities, closed_opportunities, city, state, zip, population, region) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers


#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "C:\\temp\\test\\Lundcompetencies.xlsx"
    open_file(path)