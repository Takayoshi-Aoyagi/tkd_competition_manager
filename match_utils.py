MASSOGI_MIN_PER_MATCH = 4
TUL_MIN_PER_MATCH = 3


def get_num_matches_by_tournament_level(level):
    num_matches = 0
    for i in range(level):
        num_matches += 2**i
    return num_matches


class MatchUtils:

    @classmethod
    def get_num_tournament_matches(cls, num):
        i = 0
        while True:
            num_player = 2**i
            if num == num_player:
                num_level = i
                return get_num_matches_by_tournament_level(num_level)
            if num < num_player:
                num_level = i
                num_match = get_num_matches_by_tournament_level(num_level-1)
                diff = num - 2**(i-1)
                return num_match + diff
            i += 1

    @classmethod
    def cal_time(cls, event_type, num_matches):
        if event_type == 'マッソギ':
            min_per_match = MASSOGI_MIN_PER_MATCH
        elif event_type == 'トゥル':
            min_per_match = TUL_MIN_PER_MATCH
        minutes = min_per_match * num_matches
        unit_min = 5
        units = int(minutes / unit_min)
        if minutes % unit_min:
            units += 1
        return unit_min * units

    @classmethod
    def get_timetable_items(cls, event, event_type):
        rows = []
        for c, ps in event.items():
            num_players = len(ps)
            classification = c.strip()
            if classification == '×':
                continue
            if num_players < 2:
                raise Exception(f'Invalid number of player: {classification}={num_players}')
            if num_players == 3:
                t_or_l = 'リーグ'
                num = 3
                minutes = cls.cal_time(event_type, num)
                text = f'{event_type} {classification}\n{t_or_l} {num}試合'
                rows.append([event_type, classification, num_players, num, minutes, text])
            elif num_players == 2:
                t_or_l = 'トーナメント'
                num_matches = cls.get_num_tournament_matches(num_players)
                minutes = cls.cal_time(event_type, num_matches)
                text = f'{event_type} {classification}\n決勝'
                rows.append([event_type, classification, num_players, num_matches, 10, text])
            else:
                t_or_l = 'トーナメント'
                num_matches = cls.get_num_tournament_matches(num_players)
                if event_type == 'トゥル':
                    minutes = cls.cal_time(event_type, num_matches)
                    text = f'{event_type} {classification}\n{t_or_l} {num_matches}試合'
                    rows.append([event_type, classification, num_players, num_matches, minutes, text])
                else:
                    # 決勝以外
                    num = num_matches - 1
                    minutes = cls.cal_time(event_type, num)
                    text = f'{event_type} {classification}\n{t_or_l} {num}試合'
                    rows.append([event_type, classification, num_players, num, minutes, text])
                    # 決勝
                    rows.append([
                        event_type, classification, 2, 1, 10, f'{event_type} {classification}\n決勝'])
        return rows
