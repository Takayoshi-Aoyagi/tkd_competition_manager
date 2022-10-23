import os

from aggregator import Aggregator
from excel_io import ParticipantExcelReader, DojoExcelWriter, EventExcelWriter, TournamentChartWriter


def main(outdir='participants'):
    participants = ParticipantExcelReader().execute()

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


if __name__ == '__main__':
    main()
