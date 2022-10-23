from dataclasses import dataclass
# from functools import cmp_to_key
import glob

from openpyxl import Workbook
import pandas as pd

from participant import Participant


class ParticipantExcelReader:

    def get_degree(self, val):
        if val == '段、級':
            return None
        kws = ['級', '段']
        for kw in kws:
            if kw in val:
                return val
        return None

    def get_name(self, name):
        if name in ['川口　太郎', 'わらび　花子']:
            return None
        return name

    def read_row(self, row):
        name = self.get_name(row[1])
        if pd.isnull(name):
            return None
        gender = row[3]
        degree = self.get_degree(row[4])
        if degree is None:
            return None
        dojo = row[5]
        tul = row[6]
        massogi = row[9]
        roma_name = row[13]
        kana_name = row[15]
        # print(row)
        participant = Participant(
            name=name,
            gender=gender,
            degree=degree,
            dojo=dojo,
            tul=tul,
            massogi=massogi,
            roma_name=roma_name,
            kana_name=kana_name
        )
        return participant

    def read_file(self, fpath):
        df = pd.read_excel(fpath)
        participants = []
        for index, row in df.iterrows():
            participant = self.read_row(row)
            if participant is None:
                continue
            participants.append(participant)
        return participants

    def read_files(self, input_dir):
        pattern = f'{input_dir}/*.xlsx'
        all_participants = []
        for fpath in glob.glob(pattern):
            # print(fpath)
            participants = self.read_file(fpath)
            all_participants.extend(participants)
        return all_participants

    def execute(self, input_dir='input'):
        all_participants = self.read_files(input_dir)
        # print(all_participants)
        return all_participants


@dataclass
class Result:
    event: str
    classification: str
    rank: str
    name: str


@dataclass
class ResultsExcelReader:
    fpath: str

    def get_results(self):
        COL_1ST_PRIZE = 4
        results = []
        df = pd.read_excel(self.fpath)
        for _, row in df.iterrows():
            # print(row)
            for i in range(COL_1ST_PRIZE, COL_1ST_PRIZE + 4):
                name = row[i]
                if pd.isnull(name):
                    continue
                rank = df.columns[i]
                result = Result(
                    event=row[1],
                    classification=row[3],
                    rank=rank,
                    name=name)
                results.append(result)
                print(result)
        return results

    def execute(self):
        results = self.get_results()
        return results


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
        del wb['Sheet']  # remove default sheet
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
            self.create_participants_sheet(wb, classification,
                                           category_participants)


@dataclass
class EventExcelWriter(ExcelWriter):
    event_map: map

    def create_sheets(self, wb):
        classifications = sorted(self.event_map.keys())
        for classification in classifications:
            category_participants = self.event_map[classification]
            self.create_participants_sheet(wb, classification,
                                           category_participants)
