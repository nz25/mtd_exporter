from xml.etree import ElementTree
from html.parser import HTMLParser
from itertools import product, compress, chain, zip_longest, cycle, islice
from math import ceil
from os import remove

#####################################
#
#   TOM OBJECTS
#
#####################################

class Document:

    def __init__(self, path):
        self.path = path

    def parse(self):
        tree = ElementTree.parse(self.path)
        node = tree.getroot().find('Tables')
        self.tables = [Table(n, i) for i, n in enumerate(node, start=1)]

    def __repr__(self):
        return f'Document: {self.path}'

class Table:

    def __init__(self, xml_node, index):
        self.xml_node = xml_node
        self.index = index
        self.name = xml_node.get('Name')
        self.description = xml_node.get('Description')
        self.is_populated = True if xml_node.get('IsPopulated') == 'true' else False

        # axes
        node = xml_node.find('Axes')
        self.axes = [Axis(n, self) for n in node] if node else []

        # cell items
        node = xml_node.find('CellItems')
        self.cell_items = [CellItem(n) for n in node] if node else []

        # banner
        a = [a for a in self.axes if a.name == 'Side']
        self.side_banner = Banner(self, a[0], self.cell_items) if a else None
        a = [a for a in self.axes if a.name == 'Top']
        self.top_banner = Banner(self, a[0]) if a else None

        # annotations
        self.annotations = []
        for n in xml_node.find('Annotations'):
            annotation_text = n.get('Text')
            # if annotation text starts end ends with html tags
            if '<script' in annotation_text:
                #parse html
                parser = AnnotationParser(annotation_text)
                self.annotations.append(parser.text)
            else:
                self.annotations.append(annotation_text)

        # cell values - only from the 1st layer, because
        # layers are not current supported in TOM
        node = xml_node.find('CellValues')
        self.cell_values = [
            [v for v in row.attrib.values()][1:] for row in node[0]
        ] if node else []

        # data
        self.data = self._get_data()

        # show percent signs property
        properties = xml_node.find('Properties')
        value = None
        if properties:
            for n in properties:
                if n.find('name').text == 'ShowPercentSigns':
                    value = n.find('value').text
        self.show_perc_signs = value and value == '-1'

    def _get_data(self):
        scaling_factor = len(self.cell_items)
        
        # verticalizing data
        vertical_data = [row[i::scaling_factor]
            for row in self.cell_values
            for i in range(scaling_factor)]
        
        # applying visibility masks to filter out invisible data
        if self.side_banner:
            visible_data = [row
                for row in compress(vertical_data, self.side_banner.visibility_mask)]
        else:
            visible_data = vertical_data

        if self.top_banner:
            visible_data = [
                list(compress(row, self.top_banner.visibility_mask))
                for row in visible_data
            ]

        # converting cell to numeric data types and returns
        return [[numeric(cell) for cell in row] for row in visible_data]

    @property
    def top_annotations(self):
        return self.annotations[:4]

    @property
    def bottom_annotations(self):
        return self.annotations[4:]

    def __repr__(self):
        return f'Table: {self.name}'

class Axis:
    
    def __init__(self, xml_node, host, parent=None, level=0):
        self.xml_node = xml_node
        self.host = host
        self.parent = parent
        self.level = level
        self.name = xml_node.get('Name')
        self.label = xml_node.get('Label')

        # internal
        self._expanded_elements = None
        self._expanded_element_headings = None
        self._nested_elements = None

        # sub axes
        node = xml_node.find('SubAxes')
        self.subaxes = [
            Axis(n, self.host, parent=self, level=self.level+1)
            for n in node
        ] if node else []

        # skips axis headings, which seem not to be relevant

        # elements
        node = xml_node.find('Elements')
        self.elements = [Element(n, axis=self) for n in node] if node else []

        # element headings (deduplicated)
        node = xml_node.find('ElementHeadings')
        self.element_headings = []
        if node:
            for n in node:
                element_heading = ElementHeading(n, axis=self)
                # adds only if element heading with the same name has not been added yet
                if element_heading.name not in [e.name for e in self.element_headings]:
                    self.element_headings.append(element_heading)

    @property
    def expanded_elements(self):
        if self._expanded_elements is None:
            self._expanded_elements = []
            stack = [*self.elements]
            while stack:
                current = stack[0]
                stack = stack[1:]
                self._expanded_elements.append(current)
                for sub in reversed(current.subelements):
                    stack.insert(0, sub)
        return self._expanded_elements

    @property
    def expanded_element_headings(self):
        if self._expanded_element_headings is None:
            self._expanded_element_headings = []
            stack = [*self.element_headings]
            while stack:
                current = stack[0]
                stack = stack[1:]
                self._expanded_element_headings.append(current)
                for sub in reversed(current.subelement_headings):
                    stack.insert(0, sub)
        return self._expanded_element_headings

    @property
    def nested_elements(self):
        # return list of lists, which corresponds to the
        # nested banner structure of the table
        if self._nested_elements is None:
            subaxes_elements = [e
                for c in self.subaxes
                for e in c.nested_elements]
            own_elements = [e.element
                for e in self.expanded_element_headings]
            if self.subaxes and self.elements:
                self._nested_elements = [
                    [own, *subaxes]
                    for own, subaxes in product(own_elements, subaxes_elements)]
            elif self.subaxes and not self.elements:
                self._nested_elements = subaxes_elements
            elif not self.subaxes and self.elements:
                self._nested_elements = [[e]
                    for e in own_elements]
        return self._nested_elements

    def __repr__(self):
        return f'Axis: {self.name}'

class Element:

    def __init__(self, xml_node, axis, parent=None, level=0):
        self.xml_node = xml_node
        self.axis = axis
        self.parent = parent
        self.level = level
        self.name = xml_node.get('Name')
        self.label = xml_node.get('Label')
        self.type = xml_node.get('Type')
        shown_on_table = xml_node.get('ShownOnTable')
        self.visible = shown_on_table is None or shown_on_table == 'true'
        decimals = xml_node.get('Decimals')
        self.decimals = int(decimals) if decimals else 0
        self._full_name = None

        # sub elements
        node = xml_node.find('SubElements')
        self.subelements = [
            Element(n, self.axis, parent=self, level=self.level+1)
            for n in node
        ] if node else []

    @property
    def full_name(self):
        if self._full_name is None:
            names = [self.name]
            current = self.parent
            while current:
                names.append(current.name)
                current = current.parent
            self._full_name = '.'.join(reversed(names))
        return self._full_name

    def __repr__(self):
        return f'Element: {self.full_name}'

class ElementHeading:

    def __init__(self, xml_node, axis, parent=None, level=0):
        self.xml_node = xml_node
        self.axis = axis
        self.parent = parent
        self.level = level
        self.name = xml_node.get('Name')
        self._full_name = None

        # sub element headings
        node = xml_node.find('SubElementHeadings')
        self.subelement_headings = [
            ElementHeading(n, self.axis, parent=self, level=self.level+1)
            for n in node
        ] if node else []

        # sets element
        self.element = [e
            for e in self.axis.expanded_elements 
            if e.full_name == self.full_name][0] 

    @property
    def full_name(self):
        if self._full_name is None:
            names = [self.name]
            current = self.parent
            while current:
                names.append(current.name)
                current = current.parent
            self._full_name = '.'.join(reversed(names))
        return self._full_name

    def __repr__(self):
        return f'ElementHeading: {self.full_name}'

class CellItem:

    def __init__(self, xml_node):
        self.xml_node = xml_node
        self.type = xml_node.get('Type')
        self.index = xml_node.get('Index')
        decimals = xml_node.get('Decimals')
        self.decimals = int(decimals) if decimals else 0
    def __repr__(self):
        return f'CellItem: {self.type}'

#####################################
#
#   HELPER CLASSES / FUNCTIONS
#
#####################################

class Banner:

    def __init__(self, table, axis, cell_items=None):
        self.table = table
        self.name = axis.name
        self.cell_items = cell_items
        self.scaling_factor = len(self.cell_items) if cell_items else 1

        # adding axis to elements
        elements_with_axes = [list(chain.from_iterable(
            [[e.axis, e] for e in row]))
            for row in axis.nested_elements]

        # calculates dimensions
        self.height = len(elements_with_axes) * self.scaling_factor
        self.width = max(len(c) for c in elements_with_axes) if elements_with_axes else 0

        # creates banner
        iterable_cell_items = self.cell_items if self.cell_items else [None]
        self.banner = [[BannerCell(self, cell, cell_item)
            for _, cell in zip_longest(range(self.width), row)]
            for row in elements_with_axes
            for cell_item in iterable_cell_items]

        # sets and applies visibility mask
        self.visibility_mask = [
            all(cell.visible for cell in row)
            for row in self.banner]
        self.banner = [row
            for row in compress(self.banner, self.visibility_mask)]
        self.height = len(self.banner)

        # updates labels and set first/last
        for col in range(self.width):
            for row in range(self.height):
                self._set_label(row, col)
                if self.banner[row][col].type != 'Element':
                    self._set_first_and_last(row, col)

        # creates formatting masks
        self.first_mask = [any(cell.first for cell in row) for row in self.banner]
        self.last_mask = [any(cell.last for cell in row) for row in self.banner]
        self.base_mask = [any(cell.type == 'Element' and 'Base' in cell.object.type for cell in row) for row in self.banner]
        self.cell_items_mask = list(islice(cycle(self.cell_items), self.height)) if self.cell_items else [None] * self.height
        self.last_element_mask = [[cell.object for cell in row if cell.type == 'Element'][-1] for row in self.banner]

        # propagates attributes relevant for formatting into banner cells
        for i, row in enumerate(self.banner):
            for cell in row:
                cell.first = self.first_mask[i]
                cell.last = self.last_mask[i]

        # transposes banner for Top banner
        if self.name == 'Top':
            self._transpose()

    def _transpose(self):
        self.banner = [[self.banner[row][col]
                for row in range(self.height)]
                for col in range(self.width)]
        self.height, self.width = self.width, self.height       

    def _set_label(self, row, col):
        
        cell = self.banner[row][col]
        top_cell = self.banner[row - 1][col] if row > 0 else None
        left_cell = self.banner[row][col - 1] if col > 0 else None
        top_left_cell = self.banner[row - 1][col - 1] if col > 0 and row > 0 else None

        if cell.type == 'Axis':
            if top_cell and left_cell:
                if cell.object is top_cell.object:
                    if left_cell.object is top_left_cell.object:
                        cell.label = ''
            elif top_cell and not left_cell:
                if cell.object is top_cell.object:
                    cell.label = ''
        elif cell.type == 'Element':
            if top_cell:
                if cell.object is top_cell.object:
                    cell.label = ''
     
    def _set_first_and_last(self, row, col):

        cell = self.banner[row][col]
        top_cell = self.banner[row - 1][col] if row > 0 else None
        bottom_cell = self.banner[row + 1][col] if row < self.height - 1 else None
        left_cell = self.banner[row][col - 1] if col > 0 else None
        top_left_cell = self.banner[row - 1][col - 1] if col > 0 and row > 0 else None
                   
        if not top_cell:
            cell.first = True
        if not bottom_cell:
            cell.last = True
        
        if top_cell and left_cell:
            if cell.object is not top_cell.object:
                cell.first = True
                top_cell.last = True     
            if left_cell.object is not top_left_cell.object and cell.type == 'Axis':
                cell.first = True
                top_cell.last = True     
        elif top_cell and not left_cell:
            if cell.object is not top_cell.object:
                cell.first = True
                top_cell.last = True

    def __repr__(self):
        return f'Banner: {self.name}'

class BannerCell:

    def __init__(self, banner, obj=None, cell_item=None):
        self.banner = banner
        self.object = obj
        self.cell_item = cell_item
        self.type = type(obj).__name__ if obj else ''
        self.label = obj.label if obj else ''
        self.element_type = obj.type if self.type == 'Element' else ''
        self.element_level = obj.level if self.type == 'Element' else 0
        self.axis_level = obj.level - 1 if self.type == 'Axis' else 0
        self.cell_item_type = cell_item.type if cell_item else ''
        self.visible = obj.visible if self.type == 'Element' else True
        self.first = False
        self.last = False

    def __repr__(self):
        return f'BannerCell: {self.label}'

class AnnotationParser(HTMLParser):

    def __init__(self, html):
        HTMLParser.__init__(self)
        self.data = []
        self.feed(html)

    def handle_data(self, data):
        self.data.append(data)

    @property
    def text(self):
        return self.data[-1] if self.data else ''

class Partitioner:

    def __init__(self, mtd_path, number_of_files=1):

        self.master_path = mtd_path
        self.number_of_files = number_of_files

        # reads mtd as text file
        with open(self.master_path, mode='r',encoding='utf-8') as f:
            mtd_content = f.read()

        # splits variable in header, footer and tables section
        # based on <Tables> tag
        self.header_section, tables_with_footer = mtd_content.split('<Tables>')    
        self.tables_section, self.footer_section = tables_with_footer.split('</Tables>')

        # splits table section into individual tables
        split_by = '<Table Name="'
        self.xml_tables = [f'{split_by}{t}' for t in self.tables_section.split(split_by)[1:]]
        self.number_of_tables = len(self.xml_tables)

        # ensures that xml and text parsing lead to the same number of tables
        tree = ElementTree.parse(self.master_path)
        node = tree.getroot().find('Tables')
        assert self.number_of_tables == len(node)

    def split(self):

        # number of tables per files
        tables_per_files = ceil(len(self.xml_tables)/self.number_of_files)

        # splits table list into sublists
        files = [self.xml_tables[i:i+tables_per_files]
            for i in range(0, self.number_of_tables, tables_per_files)]

        # determines file names in format: master_name_x.mtd
        file_names = [f"{self.master_path.replace('.mtd', '')}_{i}.mtd"
            for i in range(self.number_of_files)]

        for file_name, tables in zip(file_names, files):
            # joins header, table and footer section to produce text content
            content = f'{self.header_section}<Tables>{"".join(tables)}</Tables>{self.footer_section}'
            # writes splitted table names
            with open(file_name, mode='w', encoding='utf-8') as f:
                f.write(content)

        return file_names

    @classmethod
    def join(cls, file_paths, master_path, clean_up=False):

        # parses files and tables
        parsed_files = [Partitioner(f) for f in file_paths]
        tables = [t for f in parsed_files for t in f.xml_tables]

        # uses header and footer sections from the first file
        master_file = parsed_files[0]
        content = f'{master_file.header_section}<Tables>{"".join(tables)}</Tables>{master_file.footer_section}'

        # writes joined file
        with open(master_path, mode='w', encoding='utf-8') as f:
            f.write(content)

        # delete all partitioned files
        if clean_up:
            for f in file_paths:
                remove(f)


def numeric(string):
    '''Parses supplied string and returns either integer or float
    or original string if conversion is not possible.'''
    if string.isdigit():
        return int(string)
    elif string.replace('.', '', 1).replace(',', '', 1).isdigit():
        return float(string.replace(',', '.', 1))
    elif string.replace('.', '', 1).replace(',', '', 1).replace('%', '', 1).isdigit():
        return float(string.replace(',', '.', 1).replace('%', '', 1)) / 100
    else:
        return string        