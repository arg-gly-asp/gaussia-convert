import logging
import os
import re
from datetime import datetime

# GBX683P1_1 -> model 683, passage 1, attempt 1
# GBX683P1   -> attempt defaults to 1 (also when there's a bare trailing "_")
GBX_PATTERN = re.compile(r'^GBX(\d+)P(\d+)(?:_(\d*))?')


class TEST_MODEL_FILE:
    """Tracks the identifying info for one ingested Excel file.

    model        -> GBX model number, e.g. '683' from 'GBX683P1_2'
    passage      -> passage number,   e.g. '1'   from 'GBX683P1_2'
    attempt      -> int after the underscore, e.g. 2 from 'GBX683P1_2' (1 if absent)
    designation  -> single letter (A, B, C...) from group_numbers()
    first_date   -> first real datetime in column B of the sheet (injection date, day 0)
    date         -> the date of the measurement row this record came from
    values       -> the 3 replicate readings (value_1, value_2, value_3) for that mouse on that date
    """

    def __init__(self, filename, model=None, passage=None, attempt=1, designation=None, first_date=None,
                 date=None, values=None):
        if designation is not None:
            designation = str(designation).upper()
            if not re.fullmatch(r'[A-Z]', designation):
                raise ValueError(f"Designation must be a single letter, got '{designation}'")
        self.filename = filename
        self.model = model
        self.passage = passage
        self.attempt = int(attempt)
        self.designation = designation
        self.first_date = first_date
        self.date = date
        if values is not None and len(values) != 3:
            raise ValueError(f"Expected 3 replicate values, got {len(values)}")
        self.values = tuple(values) if values is not None else (None, None, None)

    @staticmethod
    def parse_sheet_title(sheet_title):
        """Return (model, passage, attempt) from a title like 'GBX683P1_2'."""
        match = GBX_PATTERN.match(sheet_title)
        if not match:
            raise ValueError(f"Sheet title '{sheet_title}' doesn't match GBX###P#[_#]")
        model, passage, attempt = match.groups()
        return model, passage, int(attempt) if attempt else 1

    @staticmethod
    def find_first_date(ws):
        """Return the first datetime in column B of an openpyxl worksheet (None if none)."""
        for cell in ws["B"]:
            if isinstance(cell.value, datetime):
                return cell.value
        return None

    @classmethod
    def from_sheet_title(cls, filename, sheet_title, designation=None, first_date=None,
                         date=None, values=None):
        """Build one from a sheet title (model, passage and attempt parsed for you)."""
        model, passage, attempt = cls.parse_sheet_title(sheet_title)
        return cls(filename, model=model, passage=passage, attempt=attempt,
                   designation=designation, first_date=first_date,
                   date=date, values=values)

    @classmethod
    def from_worksheet(cls, filename, ws, designation=None, date=None, values=None):
        """Build one straight from an openpyxl worksheet: title is parsed and
        first_date is pulled from column B."""
        return cls.from_sheet_title(filename, ws.title, designation=designation,
                                    first_date=cls.find_first_date(ws),
                                    date=date, values=values)

    def to_dict(self):
        """Handy for appending to a list of rows -> pd.DataFrame(rows)."""
        return {
            'filename': os.path.basename(self.filename),
            'model': self.model,
            'passage': self.passage,
            'attempt': self.attempt,
            'designation': self.designation,
            'first_date': self.first_date,
            'date': self.date,
            'value_1': self.values[0],
            'value_2': self.values[1],
            'value_3': self.values[2],
        }

    def __repr__(self):
        return (f"TEST_MODEL_FILE(filename={self.filename!r}, model={self.model!r}, "
                f"passage={self.passage!r}, attempt={self.attempt!r}, "
                f"designation={self.designation!r}, first_date={self.first_date!r}, "
                f"date={self.date!r}, values={self.values!r})")