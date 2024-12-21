import yaml
from collections.abc import Iterator 
class Event:
    def __init__(self,field_name,  event_data):
        self.field_name =  field_name
        self.name = event_data['name']
        self.title = event_data['title']
        self.time = event_data['time']
class Parser:
    def __init__(self, yaml_file):
        with open(yaml_file, 'r') as f:
            data = yaml.load(f, Loader=yaml.FullLoader)
            self.title = f"""{data['title']}
{data['date']}"""
            self._events = []
            for event in data['event']:
                e  = Event(event, data['event'][event])
                self._events.append(e)
    def get_events(self) -> Iterator[Event]:
        for event in self._events:
            yield event
        
    