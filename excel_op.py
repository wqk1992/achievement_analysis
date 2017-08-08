
import xlrd

    #table = data.sheets()[0]
    #table.row_values(0)
    #table.col_values(0)
    #nrows = table.nrows
    #ncols = table.ncols
class open_excel(object):
    def __init__(self, file_name = None):
        try:
            self.data = xlrd.open_workbook(file_name)
            #print(type(self.data))
        except Exception as e:
              print(str(e))

        #self.table = table
        #self.row = self.data.sheets()[table].nrows
        #self.col = self.data.sheets()[table].ncols



    # def open_excel(file = 'file_name'):
    #     try:
    #         data = xlrd.open_workbook(file)
    #         return data
    #     except Exception as e:
    #           print(str(e))

    def get_data(self, table = 0, row = 0, col = 0):
        assert isinstance(self.data, xlrd.book.Book)
        return self.data.sheets()[table].row_values(row)[col]

    def get_col(self, table = 0, col = 0):
        assert isinstance(self.data, xlrd.book.Book)
        return self.data.sheets()[table].col_values(col)

    def get_row(self, table =0, row = 0):
        assert isinstance(self.data, xlrd.book.Book)
        return self.data.sheets()[table].row_values(row)