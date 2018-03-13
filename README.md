# xlscreate
For creating simple Xls file in Python
This is simplified version for making any report in xls file which made over "xlsxwriter" with simple methods.
You can use it by following way:-

import Xcel

mySheet = Xcel('mySheetName', {
  'name' : 'Name',
  'age': 'Age',
  'belongsFrom': 'Belongs From' 
})

mySheet.createRow({
  'name': 'Vimal',
  'age': 23,
  'belongsFrom': 'Uttarakhand'
})

mySheet.createRow({
  'name': 'Aman',
  'age': 23,
  'belongsFrom': 'Punjab'
})

mySheet.closeWorkbook()