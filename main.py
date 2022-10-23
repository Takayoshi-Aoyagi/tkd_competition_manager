from aggregator import Aggregator
from excel_reader import ParticipantExcelReader
from excel_writer import PlayerExcelWriter


if __name__ == '__main__':
    participants = ParticipantExcelReader().execute()
    print(len(participants))
    massogi_map, tul_map, dojo_map = Aggregator(participants).execute()
    PlayerExcelWriter(
        participants=participants,
        massogi_map=massogi_map.get_map(),
        tul_map=tul_map.get_map(),
        dojo_map=dojo_map.get_map()
    ).execute()
