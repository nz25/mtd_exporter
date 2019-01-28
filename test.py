from mtd import Document
from xl import StandardExporter

doc = Document('test\\otto\\test.mtd')

doc.parse()

xl = StandardExporter(doc, 'test\\otto\\test.xlsx')
xl.export()
xl.save()

print('OK')