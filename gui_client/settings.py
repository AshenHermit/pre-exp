from PyQt5.QtCore import QSettings
import typing

class HandySettings():
    def __init__(self, organization: str, application: str):
        self._settings = QSettings(organization, application)
    
    def __getattr__(self, name: str) -> typing.Any:
        if name == "_settings": return super().__getattribute__(name)
        return self._settings.value(name, None)
    
    def __setattr__(self, name: str, value: typing.Any) -> None:
        if name == "_settings": super().__setattr__(name, value)
        self._settings.setValue(name, value)

storage = HandySettings('Hermit', 'Prepaid Expense Builder')