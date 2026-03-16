from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook, load_workbook

from app.excel_service import consolidate_workbooks


def test_consolidation_sums_values_from_multiple_sources() -> None:
    template = Workbook()
    template_ws = template.active
    template_ws.title = "Итог"
    template_ws["A1"] = "Наименование"
    template_ws["B1"] = "Итого"
    template_ws["A2"] = "Песок"
    template_ws["A3"] = "Щебень"

    source_1 = Workbook()
    source_1_ws = source_1.active
    source_1_ws.title = "Данные"
    source_1_ws["A1"] = "Наименование"
    source_1_ws["B1"] = "Количество"
    source_1_ws["A2"] = "Песок"
    source_1_ws["B2"] = 10
    source_1_ws["A3"] = "Щебень"
    source_1_ws["B3"] = 5

    source_2 = Workbook()
    source_2_ws = source_2.active
    source_2_ws.title = "Данные"
    source_2_ws["A1"] = "Наименование"
    source_2_ws["B1"] = "Количество"
    source_2_ws["A2"] = "Песок"
    source_2_ws["B2"] = 2
    source_2_ws["A3"] = "Щебень"
    source_2_ws["B3"] = 7

    config = {
        "template_cols_preset": {
            "default": {
                "key": {"header": "Наименование"},
                "qty": {"header": "Итого"},
            }
        },
        "source_cols_preset": {
            "default": {
                "key": {"header": "Наименование"},
                "qty": {"header": "Количество"},
            }
        },
        "pages": [
            {
                "name": "Основной лист",
                "template_sheet": "Итог",
                "source_sheet": "Данные",
                "template_preset": "default",
                "source_preset": "default",
                "match_field": "key",
                "mappings": [
                    {
                        "template_field": "qty",
                        "source_field": "qty",
                        "aggregate": "sum",
                        "default": 0,
                    }
                ],
            }
        ],
    }

    result_bytes, _, report = consolidate_workbooks(
        template_bytes=_dump_book(template),
        template_name="template.xlsx",
        sources=[
            ("source_1.xlsx", _dump_book(source_1)),
            ("source_2.xlsx", _dump_book(source_2)),
        ],
        config=config,
    )

    result_book = load_workbook(BytesIO(result_bytes))
    result_ws = result_book["Итог"]

    assert result_ws["B2"].value == 12
    assert result_ws["B3"].value == 12
    assert report["pages"][0]["matched_rows"] == 2


def _dump_book(workbook: Workbook) -> bytes:
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()
