    def make_all_data_sheet(self, results: Sequence[SampleResult]) -> str:
        rows = [
            self.row(1, [self.cell(1, 1, "Sample", 1), self.cell(1, 2, "Sample Name", 1), self.cell(1, 3, "Row", 1), self.cell(1, 4, "Col", 1), self.cell(1, 5, "MPCD", 1)])
        ]
        row_index = 2
        for sample in results:
            for r, source_row in enumerate(sample.values, 1):
                for c, value in enumerate(source_row, 1):
                    rows.append(
                        self.row(
                            row_index,
                            [
                                self.cell(row_index, 1, sample.sample_id),
                                self.cell(row_index, 2, sample_title(sample)),
                                self.cell(row_index, 3, r),
                                self.cell(row_index, 4, c),
                                self.cell(row_index, 5, value, self.style_for_color(interpolate_color(value))),
                            ],
                        )
                    )
                    row_index += 1
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:E{row_index - 1}"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16"/>'
            '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="5" width="12" customWidth="1"/></cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '</worksheet>'
        )

 
