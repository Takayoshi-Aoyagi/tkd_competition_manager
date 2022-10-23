from functools import cmp_to_key

from openpyxl import Workbook

from dataclasses import dataclass

@dataclass
class ExcelWriter:

    filename: str

    def create_participants_sheet(self, wb, title, participants):
        sheet = wb.create_sheet(title=title)
        row = 1

        headers = ['氏名', 'かな', 'Name', '性別', '級位', '道場名', 'トゥル', 'マッソギ']
        for col, header in enumerate(headers):
            sheet.cell(column=col+1, row=row, value=header)

        for p in participants:
            row += 1
            sheet.cell(column=1, row=row, value=p.name)
            sheet.cell(column=2, row=row, value=p.kana_name)
            sheet.cell(column=3, row=row, value=p.roma_name)
            sheet.cell(column=4, row=row, value=p.gender)
            sheet.cell(column=5, row=row, value=p.degree)
            sheet.cell(column=6, row=row, value=p.dojo)
            sheet.cell(column=7, row=row, value=p.tul)
            sheet.cell(column=8, row=row, value=p.massogi)            
        
    def execute(self):
        wb = Workbook()
        self.create_sheets(wb)
        del wb['Sheet'] # remove default sheet
        wb.save(filename=self.filename)


@dataclass
class DojoExcelWriter(ExcelWriter):
    participants: list
    dojo_map: map

    def create_sheets(self, wb):
        participants = []
        for _, ps in self.dojo_map.items():
            participants.extend(ps)
        self.create_participants_sheet(wb, '選手一覧', participants)

        for classification, category_participants in self.dojo_map.items():
            self.create_participants_sheet(wb, classification, category_participants)


@dataclass
class EventExcelWriter(ExcelWriter):
    event_map: map

    def create_sheets(self, wb):
        classifications = sorted(self.event_map.keys())
        for classification in classifications:
            category_participants = self.event_map[classification]
            self.create_participants_sheet(wb, classification, category_participants)
        

