from pathlib import Path

class PathReference():
    def __init__(self, path:Path=None) -> None:
        self.path:Path = path or Path()
class StringReference():
    def __init__(self, string:str=None) -> None:
        self.string:str = string or ""

    def make_widget_updater(self, widget):
        def on_update():
            self.string = widget.text()
        on_update()
        return on_update