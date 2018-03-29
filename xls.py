import xlsxwriter


class Xcel:
    def __init__(self, name, schema):
      self.name = 'reports/' + name + '.xlsx'
      self.schema = schema
      self.workbook = xlsxwriter.Workbook(self.name)
      self.style = {
          'bold': self.workbook.add_format({'bold': True})
      }
      self.ALPHABETS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
      self.worksheet = self.workbook.add_worksheet()
      self.row = 0
      self.columns = []
      col = 0
      i = 0
      for colVal, colName in self.schema.items():
          colIndex = self.ALPHABETS[i] + '1'
          self.columns.append(colVal)
          self.worksheet.write(colIndex, colName, self.style['bold'])
          i += 1
      i = 0
      self.row += 1

    def createRow(self, data):
      row = self.row
      col = 0
      for colVal in self.columns:
        try:
          colVal = str(colVal)
          value = str(data[colVal])
          self.worksheet.write(row, col, value)
        except:
          pass
        col += 1
      self.row += 1
      return self

    def closeWorkbook(self):
      print('saving :- ', self.name)
      self.workbook.close()
