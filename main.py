from aggregator import Aggregator
from excel_reader import ParticipantExcelReader
from excel_writer import DojoExcelWriter, EventExcelWriter


def main():
    participants = ParticipantExcelReader().execute()

    massogi_map, tul_map, dojo_map = Aggregator(participants).execute()
    DojoExcelWriter(
        filename='道場別選手一覧.xlsx',
        participants=participants,
        dojo_map=dojo_map.get_map()
    ).execute()
    EventExcelWriter(
        filename='マッソギ選手一覧.xlsx',
        event_map=massogi_map.get_map()
    ).execute()
    EventExcelWriter(
        filename='トゥル選手一覧.xlsx',
        event_map=tul_map.get_map()
    ).execute()


if __name__ == '__main__':
    main()
