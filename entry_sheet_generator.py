from argparse import ArgumentParser
import os

from aggregator import Aggregator
from excel_io import (
    ParticipantExcelReader,
    DojoExcelWriter,
    EventExcelWriter,
    TournamentChartWriter,
    TimetableWriter,
    WinnerListWriter
)
from merger import Merger


def dump(class_participant_map):
    print("========================")
    print(class_participant_map.name)
    for key, items in class_participant_map._map.items():
        print(key, items)


def main(args):
    outdir = args.outdir
    participants = ParticipantExcelReader(
        vacant_to_withdraw=args.vacant_to_withdraw
    ).execute()
    participants = Merger(participants).execute()

    massogi_map, tul_map, dojo_map = Aggregator(participants).execute()
    DojoExcelWriter(
        filename=os.path.join(outdir, '道場別選手一覧.xlsx'),
        participants=participants,
        dojo_map=dojo_map.get_map()
    ).execute()
    EventExcelWriter(
        filename=os.path.join(outdir, 'マッソギ選手一覧.xlsx'),
        event_map=massogi_map.get_map()
    ).execute()
    EventExcelWriter(
        filename=os.path.join(outdir, 'トゥル選手一覧.xlsx'),
        event_map=tul_map.get_map()
    ).execute()

    TournamentChartWriter(
        filename=os.path.join(outdir, 'マッソギ対戦表.xlsx'),
        event_map=massogi_map.get_map()
    ).execute()
    TournamentChartWriter(
        filename=os.path.join(outdir, 'トゥル対戦表.xlsx'),
        event_map=tul_map.get_map()
    ).execute()

    TimetableWriter(
        filename=os.path.join(outdir, 'タイムテーブル.xlsx'),
        massogi=massogi_map.get_map(),
        tul=tul_map.get_map()
    ).execute()

    WinnerListWriter(
        filename=os.path.join(outdir, 'winners_blank.xlsx'),
        massogi=massogi_map.get_map(),
        tul=tul_map.get_map()
    ).execute()


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument('--outdir', default='participants')
    parser.add_argument('--vacant-to-withdraw', action='store_true')
    args = parser.parse_args()
    main(args)
