from dataclasses import dataclass
import json
import os

from participant import Participant


@dataclass
class Merger:
    participants: list[Participant]
    config_path: str = 'merge.json'

    def get_config(self):
        with open(self.config_path) as fp:
            return json.load(fp)

    def merge(self, config, p):
        for event in config.keys():  # tul or massogi
            mapping = config[event]
            curr_classes = mapping.keys()
            curr_class = getattr(p, event)
            if curr_class not in curr_classes:
                continue
            merged_class = mapping[curr_class]
            print(p)
            setattr(p, event, merged_class)
            print(p)

    def execute(self):
        if not os.path.exists(self.config_path):
            return self.participants
        config = self.get_config()
        for p in self.participants:
            self.merge(config, p)
        return self.participants
