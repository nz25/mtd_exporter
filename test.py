from mtd import Document
from xl import StandardExporter

doc = Document('test.mtd')

doc.parse()

xl = StandardExporter(doc, 'test.xlsx')
xl.export()
xl.save()

print('OK')