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

        changed_l = []
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
                    changed_l.append(f"[{sheet.title}:row={row_no}] {colored_text}")

        if changed_l:
            self._output_file('./_outputs/changed.log', changed_l)

    def _output_file(self, output_f: str, contents_l: list):
        if not contents_l:
            return ''

        out_dir = os.path.dirname(output_f)
        if not os.path.exists(out_dir):
            os.mkdir(out_dir)

        with open(output_f, 'w', encoding="utf-8") as f:
            f.writelines('\n'.join(contents_l))

        self._logger.debug(f"output csv: {output_f}")
