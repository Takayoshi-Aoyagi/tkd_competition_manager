class ClassParticipantMap:

    def __init__(self, name):
        self.name = name
        self._map = {}

    def append(self, classification, participant):
        if classification not in self._map:
            self._map[classification] = []
        self._map[classification].append(participant)

    def get_map(self):
        return self._map


class Aggregator:

    def __init__(self, participants):
        self.participants = participants
        self.massogi_map = ClassParticipantMap('Massogi')
        self.tul_map = ClassParticipantMap('Tul')
        self.dojo_map = ClassParticipantMap('Dojo')

    def execute(self):
        for p in self.participants:
            self.massogi_map.append(p.massogi, p)
            self.tul_map.append(p.tul, p)
            self.dojo_map.append(p.dojo, p)
        return self.massogi_map, self.tul_map, self.dojo_map
            
