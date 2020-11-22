# -*- coding: utf-8 -*-

import logging
import os
import openpyxl


class ExecelChangeDetector:

    def __init__(self):
        self._logger = logging.getLogger(__name__)

    def detect_font_changed_row(self, excel_file):
        """
        Detect file row which font color are set by rgb.
        """
        wb = openpyxl.load_workbook(excel_file, read_only=True)
        for sheet in wb:
            for row in sheet.rows:
                colored_text = ""
                row_no = ""
                for cell in row:
                    if cell.value and cell.font and cell.font.color and cell.font.color.type == 'rgb':
                        colored_text += "," if colored_text else ""
                        colored_text += f"{cell.value}"
                        row_no = cell.row
                if colored_text:
                    self._logger.debug(f"[{sheet.title}:row={row_no}] {colored_text}")
