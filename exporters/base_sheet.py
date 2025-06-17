# base_sheet.py

class BaseSheetExporter:
    def __init__(self, df, writer):
        self.df = df
        self.writer = writer
