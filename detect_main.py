# -*- coding: utf-8 -*-

import argparse
import logging
from detect_excel_changed.excel_change_detecter import ExecelChangeDetector

if __name__ == '__main__':

    _logger = logging.getLogger('detect_excel_changed')
    _logger.setLevel(logging.DEBUG)
    # handler
    handler = logging.StreamHandler()
    handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(filename)-12s: %(message)s')
    handler.setFormatter(formatter)
    _logger.addHandler(handler)

    try:
        _logger.debug("start~~")

        filepath = r".\tmp_data\000662421.xlsx"

        detector = ExecelChangeDetector()
        detector.detect_font_changed_row(filepath)

        _logger.debug("fin.")

    except Exception as e:
        _logger.exception(e)
