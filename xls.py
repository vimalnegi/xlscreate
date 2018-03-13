import xlsxwriter


class Xcel:
    def __init__(self, name, schema):
      self.name = name + '.xlsx'
      self.schema = schema
      self.workbook = xlsxwriter.Workbook(self.name)
      self.worksheet = self.workbook.add_worksheet()
      self.row = 0
      self.columns = []
      col = 0
      i = 0
      for colVal, colName in self.schema.items():
          self.columns.append(colVal)
          self.worksheet.write(self.row, col + i, colName)
          i += 1
      i = 0
      self.row += 1

    def createRow(self, data):
      row = self.row
      col = 0
      for colVal in self.columns:
          colVal = str(colVal)
          value = str(data[colVal])
          self.worksheet.write(row, col, value)
          col += 1
      self.row += 1
      return self

    def closeWorkbook(self):
      print('saving :- ', self.name)
      self.workbook.close()
