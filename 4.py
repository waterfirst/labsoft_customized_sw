   def make_summary_sheet(self, results: Sequence[SampleResult], folder: Path, layout: str, source_mode: str) -> str:
        data: List[List[object]] = [
            ["MPCD Excel Program"],
            ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            ["Folder", str(folder)],
            ["Source mode", source_mode],
            ["Layout", layout],
            ["Color scale", f"{COLOR_MIN_VALUE:g}=#{COLOR_MIN}, {COLOR_MID_VALUE:g}=#{COLOR_MID}, {COLOR_MAX_VALUE:g}=#{COLOR_MAX}"],
            [],
            ["Sample Name", "Source", "Count", "Mean", "Std", "Min", "Max"],
        ]
        for sample in results:
            summary = summarize_values(sample.values)
            data.append(
                [
                    sample_title(sample),
                    sample.source,
                    summary["count"],
                    summary["mean"],
                    summary["std"],
                    summary["min"],
                    summary["max"],
                ]
            )

        rows = []
        for row_index, row_values in enumerate(data, 1):
            cells = []
            for col_index, value in enumerate(row_values, 1):
                cells.append(self.cell(row_index, col_index, value, 1 if row_index in (1, 8) else None))
            rows.append(self.row(row_index, cells))
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:G{len(data)}"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16"/>'
            '<cols><col min="1" max="7" width="18" customWidth="1"/></cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '</worksheet>'
        )

    def make_graph_sheet(self, results: Sequence[SampleResult]) -> str:
        def mean_line(values: Iterable[Optional[float]]) -> Optional[float]:
            data = [value for value in values if value is not None and math.isfinite(value)]
            return statistics.mean(data) if data else None

        rows: List[str] = []
        row_index = 1
        max_col = 37
        for sample in results:
            summary = summarize_values(sample.values)
            rows.append(self.row(row_index, [self.cell(row_index, 1, f"{sample_title(sample)} Graph", 1)]))
            row_index += 1
            rows.append(
                self.row(
                    row_index,
                    [
                        self.cell(row_index, 1, "Mean", 1),
                        self.cell(row_index, 2, summary["mean"], 3),
                        self.cell(row_index, 3, "Min", 1),
                        self.cell(row_index, 4, summary["min"], 3),
                        self.cell(row_index, 5, "Max", 1),
                        self.cell(row_index, 6, summary["max"], 3),
                    ],
                )
            )
            row_index += 2

            row_means = [mean_line(row) for row in sample.values]
            rows.append(self.row(row_index, [self.cell(row_index, 1, "Row Profile", 1)]))
            row_index += 1
            header = [self.cell(row_index, 1, "Row", 2)]
            for idx in range(len(row_means)):
                header.append(self.cell(row_index, 2 + idx, idx + 1, 2))
            rows.append(self.row(row_index, header))
            row_index += 1
            value_cells = [self.cell(row_index, 1, "MPCD", 2)]
            for idx, value in enumerate(row_means):
                value_cells.append(self.cell(row_index, 2 + idx, value, self.style_for_color(interpolate_color(value))))
            rows.append(self.row(row_index, value_cells))
            row_index += 3

            col_count = sample.cols
            col_means = []
            for col in range(col_count):
                col_means.append(mean_line(row[col] for row in sample.values if col < len(row)))
            rows.append(self.row(row_index, [self.cell(row_index, 1, "Column Profile", 1)]))
            row_index += 1
            header = [self.cell(row_index, 1, "Column", 2)]
            for idx in range(col_count):
                header.append(self.cell(row_index, 2 + idx, f"C{idx + 1}", 2))
            rows.append(self.row(row_index, header))
            row_index += 1
            value_cells = [self.cell(row_index, 1, "MPCD", 2)]
            for idx, value in enumerate(col_means):
                value_cells.append(self.cell(row_index, 2 + idx, value, self.style_for_color(interpolate_color(value))))
            rows.append(self.row(row_index, value_cells))
            row_index += 3

            side_summary = sample.side_summary or compute_side_summary(sample.values)
            rows.append(self.row(row_index, [self.cell(row_index, 1, "Summary Graph", 1)]))
            row_index += 1
            cells = [self.cell(row_index, 1, "Item", 2)]
            for idx, (label, _) in enumerate(side_summary):
                cells.append(self.cell(row_index, 2 + idx, label, 2))
            rows.append(self.row(row_index, cells))
            row_index += 1
            cells = [self.cell(row_index, 1, "MPCD", 2)]
            for idx, (_, value) in enumerate(side_summary):
                cells.append(self.cell(row_index, 2 + idx, value, self.style_for_color(interpolate_color(value))))
            rows.append(self.row(row_index, cells))
            row_index += 4

        cols = [
            '<col min="1" max="1" width="14" customWidth="1"/>',
            f'<col min="2" max="{max_col}" width="7" customWidth="1"/>',
        ]
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:{cell_ref(max(1, row_index - 1), max_col)}"/>'
            '<sheetViews><sheetView workbookViewId="0" zoomScale="85" zoomScaleNormal="85"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16"/>'
            f'<cols>{"".join(cols)}</cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '</worksheet>'
        )

    def styles_xml(self) -> str:
        fills = [
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FF404040"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>',
        ]
        for color in self.colors:
            fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{color}"/><bgColor indexed="64"/></patternFill></fill>')

        xfs = [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>',
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>',
            '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>',
        ]
        for fill_id in range(4, 4 + len(self.colors)):
            xfs.append(
                f'<xf numFmtId="164" fontId="0" fillId="{fill_id}" borderId="1" xfId="0" '
                'applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>'
            )

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts>'
            '<fonts count="2">'
            '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
            '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
            '</fonts>'
            f'<fills count="{len(fills)}">{"".join(fills)}</fills>'
            '<borders count="2">'
            '<border><left/><right/><top/><bottom/><diagonal/></border>'
            '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'
            '</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
            '</styleSheet>'
        )


def write_xlsx(
    output_path: Path,
    results: Sequence[SampleResult],
    folder: Path,
    layout: str,
    source_mode: str,
    time_groups: Sequence[str],
) -> None:
    builder = WorkbookBuilder()
    new_result_sheet = builder.make_new_result_sheet(results)
    map_sheet = builder.make_map_sheet(results, layout, time_groups)
    all_data_sheet = builder.make_all_data_sheet(results)
    summary_sheet = builder.make_summary_sheet(results, folder, layout, source_mode)
    graph_sheet = builder.make_graph_sheet(results)
    styles = builder.styles_xml()
    chart_overrides = "".join(
        f'<Override PartName="/xl/charts/chart{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        for index, _ in enumerate(results, 1)
    )
    drawing_overrides = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' if results else ""

    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            f'{drawing_overrides}'
            f'{chart_overrides}'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>'
            '<sheet name="새로운 결과1" sheetId="1" r:id="rId1"/>'
            '<sheet name="MPCD Map" sheetId="2" r:id="rId2"/>'
            '<sheet name="All Data" sheetId="3" r:id="rId3"/>'
            '<sheet name="Summary" sheetId="4" r:id="rId4"/>'
            '<sheet name="Graph" sheetId="5" r:id="rId5"/>'
            '</sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
            '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>'
            '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>'
            '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ),
        "xl/worksheets/_rels/sheet1.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            '</Relationships>'
        ),
        "xl/worksheets/sheet1.xml": new_result_sheet,
        "xl/worksheets/sheet2.xml": map_sheet,
        "xl/worksheets/sheet3.xml": all_data_sheet,
        "xl/worksheets/sheet4.xml": summary_sheet,
        "xl/worksheets/sheet5.xml": graph_sheet,
        "xl/styles.xml": styles,
        "docProps/core.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>mpcd_excel_program.py</dc:creator>'
            '<cp:lastModifiedBy>mpcd_excel_program.py</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}</dcterms:created>'
            '</cp:coreProperties>'
        ),
        "docProps/app.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Python</Application></Properties>'
        ),
    }
    if results:
        files["xl/drawings/drawing1.xml"] = make_sheet1_drawing_xml(results)
        files["xl/drawings/_rels/drawing1.xml.rels"] = make_sheet1_drawing_rels(results)
        for index, sample in enumerate(results, 1):
            base_row = 1 + (index - 1) * (65 + 3)
            header_row = base_row + 40
            first_data_row = header_row + 1
            files[f"xl/charts/chart{index}.xml"] = make_line_chart_xml("새로운 결과1", f"{sample_title(sample)} Profile", header_row, first_data_row)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)


def write_csv_output(csv_path: Path, results: Sequence[SampleResult]) -> None:
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Sample", "Sample Name", "Row", "Col", "MPCD"])
        for sample in results:
            for row_index, row in enumerate(sample.values, 1):
                for col_index, value in enumerate(row, 1):
                    writer.writerow([sample.sample_id, sample_title(sample), row_index, col_index, "" if value is None else f"{value:.10g}"])


def sheet_formula(sheet_name: str, cell_range: str) -> str:
    return f"'{sheet_name.replace(chr(39), chr(39) + chr(39))}'!{cell_range}"


def make_line_chart_xml(sheet_name: str, chart_title: str, header_row: int, first_data_row: int) -> str:
    cat_formula = sheet_formula(sheet_name, f"$D${header_row}:$R${header_row}")
    colors = ("4472C4", "ED7D31", "70AD47")
    series_xml = []
    for index, color in enumerate(colors):
        row = first_data_row + index
        label_formula = sheet_formula(sheet_name, f"$C${row}")
        value_formula = sheet_formula(sheet_name, f"$D${row}:$R${row}")
        series_xml.append(
            f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/>'
            f'<c:tx><c:strRef><c:f>{escape(label_formula)}</c:f></c:strRef></c:tx>'
            f'<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>'
            '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>'
            f'<c:cat><c:strRef><c:f>{escape(cat_formula)}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{escape(value_formula)}</c:f></c:numRef></c:val>'
            '<c:smooth val="0"/></c:ser>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:lang val="ko-KR"/><c:roundedCorners val="0"/>'
        '<c:chart>'
        f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{escape(chart_title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>'
        '<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        f'{"".join(series_xml)}'
        '<c:axId val="50010001"/><c:axId val="50010002"/></c:lineChart>'
        '<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
        '<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="0.0" sourceLinked="0"/><c:majorTickMark val="out"/>'
        '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010001"/><c:crosses val="autoZero"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart>'
        '</c:chartSpace>'
    )


def make_sheet1_drawing_xml(results: Sequence[SampleResult]) -> str:
    anchors = []
    block_height = 65
    block_gap = 3
    for index, sample in enumerate(results, 1):
        base_row = 1 + (index - 1) * (block_height + block_gap)
        from_row = base_row + 46
        to_row = base_row + 64
        anchors.append(
            '<xdr:twoCellAnchor>'
            f'<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>18</xdr:col><xdr:colOff>325290</xdr:colOff><xdr:row>{to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro="">'
            f'<xdr:nvGraphicFramePr><xdr:cNvPr id="{index + 1}" name="Chart {index}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{index}"/>'
            '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'{"".join(anchors)}</xdr:wsDr>'
    )


def make_sheet1_drawing_rels(results: Sequence[SampleResult]) -> str:
    relationships = [
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{index}.xml"/>'
        for index, _ in enumerate(results, 1)
    ]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(relationships)}</Relationships>'
    )


def filter_results_by_sample(results: Sequence[SampleResult], sample_selection: str) -> List[SampleResult]:
    selected_ids = parse_sample_selection_ids(sample_selection)
    if selected_ids is None:
        return list(results)
    selected_set = set(selected_ids)
    filtered = [sample for sample in results if sample.sample_id in selected_set]
    if not filtered:
        sample_text = ", ".join(f"Sample {sample_id}" for sample_id in selected_ids)
        raise ValueError(f"선택한 {sample_text} 결과를 찾지 못했습니다.")
    return filtered


def run_export(
    folder: Path,
    output_path: Path,
    layout: str = "excel_style",
    source_mode: str = "raw_txt",
    reference_excel: Optional[Path] = None,
    time_groups: Sequence[str] = (),
    sample_selection: str = "전체",
    panel_roi: float = 0.98,
    progress: Optional[callable] = None,
    manual_pcf_files: Optional[Sequence[Path]] = None,
) -> List[SampleResult]:
    if progress:
        if source_mode == "reference_excel":
            progress("정리 Excel에서 MPCD 결과를 읽는 중...")
        elif os.environ.get("MPCD_FORCE_PCF", "0") == "1":
            progress("PCF에서 Lxy/WST를 추출한 뒤 MPCD 계산 중...")
        else:
            progress("Raw txt에서 MPCD 계산 중...")
    if source_mode == "reference_excel":
        results = filter_results_by_sample(read_reference_excel(folder, reference_excel), sample_selection)
    else:
        results = read_raw_results(folder, panel_roi, progress, sample_selection, manual_pcf_files)
    if progress:
        progress(f"{len(results)}개 시료 Excel / CSV 저장 중...")
    write_xlsx(output_path, results, folder, layout, source_mode, time_groups)
    write_csv_output(output_path.with_suffix(".csv"), results)
    if progress:
        progress(f"완료: {output_path}")
    return results


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("15x35 MPCD 분석")
        self.geometry("1180x820")
        self.minsize(1040, 720)
        self.resizable(True, True)
        self.configure(bg="#EFF3F8")
        self.after(50, self.maximize_window)

        cwd = Path.cwd()
        self.settings = load_app_settings()
        self.panel_roi_var = tk.StringVar(value="0.98")
        self.folder_var = tk.StringVar(value=str(cwd))
        self.labsoft_path_var = tk.StringVar(value=self.settings.get("labsoft_path", default_labsoft_path()))
        self.force_pcf_var = tk.BooleanVar(value=False)
        self.reference_var = tk.StringVar(value=str(suggest_reference_excel(cwd)))
        self.output_var = tk.StringVar(value=str(automatic_output_path(cwd)))
        self.output_name_var = tk.StringVar(value=automatic_output_path(cwd).name)
        self.layout_var = tk.StringVar(value="excel_style")
        self.status_var = tk.StringVar(value="대기 중")
        self.last_output: Optional[Path] = None
        self.last_csv: Optional[Path] = None
        self.last_results: List[SampleResult] = []
        self.manual_pcf_files: List[Path] = []
        self.cancel_event = threading.Event()
        self.folder_value_label: Optional[tk.Label] = None
        self.labsoft_value_label: Optional[tk.Entry] = None
        self.reference_value_label: Optional[tk.Label] = None
        self.output_value_label: Optional[tk.Label] = None
        self.result_card: Optional[tk.Frame] = None
        self.result_title_label: Optional[tk.Label] = None
        self.result_meta_label: Optional[tk.Label] = None
        self.result_detail_label: Optional[tk.Label] = None
        self.create_button: Optional[tk.Button] = None
        self.stop_button: Optional[tk.Button] = None
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.sample_listbox: Optional[tk.Listbox] = None
        self.manual_pcf_label: Optional[tk.Label] = None
        self.result_sample_var = tk.StringVar(value="")
        self.result_sample_combo: Optional[ttk.Combobox] = None
        self.result_table_frame: Optional[tk.Frame] = None
        self.result_text: Optional[tk.Text] = None
        self._build_ui()
        self.refresh_sample_options()

    def _build_ui(self) -> None:
        self.option_add("*Font", ("Malgun Gothic", 10))
        style = ttk.Style(self)
        style.configure("MPCD.Horizontal.TProgressbar", troughcolor="#E8EEF8", background="#1769D1")

        shell = tk.Frame(self, bg="#EFF3F8")
        shell.pack(fill="both", expand=True, padx=40, pady=18)

        tk.Label(
            shell,
            text="15x35 MPCD 분석",
            font=("Malgun Gothic", 22, "bold"),
            fg="#071B3D",
            bg="#EFF3F8",
        ).pack(anchor="w", pady=(0, 16))

        input_card = self.card(shell)
        input_card.pack(fill="x", pady=(0, 18))
        for col, weight in enumerate((0, 2, 0, 2, 0, 1)):
            input_card.grid_columnconfigure(col, weight=weight)

        def label(parent: tk.Misc, text: str, row: int, column: int) -> None:
            tk.Label(parent, text=text, font=("Malgun Gothic", 10, "bold"), fg="#0F1E33", bg="#FFFFFF").grid(
                row=row, column=column, sticky="w", padx=(2, 10), pady=8
            )

        def entry(parent: tk.Misc, variable: tk.StringVar, row: int, column: int, width: int = 26) -> tk.Entry:
            field = tk.Entry(
                parent,
                textvariable=variable,
                relief="solid",
                bd=1,
                highlightthickness=0,
                bg="#FFFFFF",
                fg="#071B3D",
                insertbackground="#071B3D",
                font=("Malgun Gothic", 10),
                width=width,
            )
            field.grid(row=row, column=column, sticky="ew", padx=(0, 10), pady=8, ipady=6)
            return field

        def path_row(row: int, title: str, value_var: tk.StringVar, choose_text: str, choose_command: callable, open_command: callable) -> tk.Label:
            label(input_card, title, row, 0)
            box = tk.Frame(input_card, bg="#FFFFFF")
            box.grid(row=row, column=1, columnspan=5, sticky="ew", pady=5)
            box.grid_columnconfigure(1, weight=1)
            self.button(box, choose_text, choose_command, fill="#F5F9FF", fg="#071B3D").grid(row=0, column=0, sticky="w", padx=(0, 8), ipady=2)
            value_label = tk.Label(
                box,
                text=self.display_path(Path(value_var.get())),
                fg="#071B3D",
                bg="#FFFFFF",
                anchor="w",
                relief="solid",
                bd=1,
                padx=10,
                pady=7,
            )
            value_label.grid(row=0, column=1, sticky="ew")
            self.button(box, "열기", open_command, fill="#EEF5FF", fg="#1257B2").grid(row=0, column=2, sticky="e", padx=(8, 0), ipady=2)
            return value_label

        def path_entry_row(row: int, title: str, value_var: tk.StringVar, choose_text: str, choose_command: callable, open_command: callable) -> tk.Entry:
            label(input_card, title, row, 0)
            box = tk.Frame(input_card, bg="#FFFFFF")
            box.grid(row=row, column=1, columnspan=5, sticky="ew", pady=5)
            box.grid_columnconfigure(0, weight=1)
            field = tk.Entry(
                box,
                textvariable=value_var,
                relief="solid",
                bd=1,
                highlightthickness=0,
                bg="#FFFFFF",
                fg="#071B3D",
                insertbackground="#071B3D",
                font=("Malgun Gothic", 10),
            )
            field.grid(row=0, column=0, sticky="ew", padx=(0, 8), ipady=7)
            self.button(box, choose_text, choose_command, fill="#F5F9FF", fg="#071B3D").grid(row=0, column=1, sticky="e", padx=(0, 8), ipady=2)
            self.button(box, "열기", open_command, fill="#EEF5FF", fg="#1257B2").grid(row=0, column=2, sticky="e", ipady=2)
            return field

        label(input_card, "Panel ROI", 0, 0)
        entry(input_card, self.panel_roi_var, 0, 1, width=10)

        self.folder_value_label = path_row(
            1,
            "Data 입력",
            self.folder_var,
            "폴더 선택",
            self.choose_folder,
            lambda: self.open_path(Path(self.folder_var.get())),
        )
        self.labsoft_value_label = path_entry_row(
            2,
            "LMK LabSoft",
            self.labsoft_path_var,
            "찾아보기",
            self.choose_labsoft_path,
            lambda: self.open_path(Path(self.labsoft_path_var.get())),
        )

        label(input_card, "분석 시료 선택", 3, 0)
        option_row = tk.Frame(input_card, bg="#FFFFFF")
        option_row.grid(row=3, column=1, columnspan=5, sticky="ew", pady=(8, 4))
        option_row.grid_columnconfigure(1, weight=1)
        tk.Label(option_row, text="시료번호", fg="#52617A", bg="#FFFFFF").grid(row=0, column=0, sticky="nw", pady=(6, 0))
        sample_box = tk.Frame(option_row, bg="#FFFFFF")
        sample_box.grid(row=0, column=1, columnspan=5, sticky="ew", padx=(8, 0))
        sample_box.grid_columnconfigure(0, weight=1)
        self.sample_listbox = tk.Listbox(
            sample_box,
            selectmode="extended",
            exportselection=False,
            height=5,
            relief="solid",
            bd=1,
            bg="#FFFFFF",
            fg="#071B3D",
            activestyle="none",
            font=("Malgun Gothic", 10),
        )
        self.sample_listbox.grid(row=0, column=0, sticky="ew")
        sample_scroll = ttk.Scrollbar(sample_box, orient="vertical", command=self.sample_listbox.yview)
        sample_scroll.grid(row=0, column=1, sticky="ns")
        self.sample_listbox.configure(yscrollcommand=sample_scroll.set)
        sample_actions = tk.Frame(sample_box, bg="#FFFFFF")
        sample_actions.grid(row=0, column=2, sticky="ns", padx=(8, 0))
        self.button(sample_actions, "전체 선택", self.select_all_samples, fill="#EEF5FF", fg="#1257B2").pack(fill="x", pady=(0, 6))
        self.button(sample_actions, "선택 해제", self.clear_sample_selection, fill="#F5F9FF", fg="#1257B2").pack(fill="x")
        self.button(sample_actions, "TXT 파일 선택", self.choose_txt_files, fill="#EEF5FF", fg="#1257B2").pack(fill="x", pady=(8, 6))
        self.button(sample_actions, "PCF 파일 선택", self.choose_pcf_files, fill="#EEF5FF", fg="#1257B2").pack(fill="x", pady=(8, 6))
        self.button(sample_actions, "PCF 선택 해제", self.clear_pcf_files, fill="#F5F9FF", fg="#1257B2").pack(fill="x")
        layout_buttons = tk.Frame(option_row, bg="#FFFFFF")
        layout_buttons.grid(row=1, column=1, columnspan=5, sticky="ew", padx=(8, 0), pady=(10, 0))
        self.add_layout_button(layout_buttons, "정리 Excel형", "excel_style")
        self.add_layout_button(layout_buttons, "전체 아래로", "vertical")
        self.manual_pcf_label = tk.Label(option_row, text="", fg="#536789", bg="#FFFFFF", anchor="w", justify="left")
        self.manual_pcf_label.grid(row=2, column=1, columnspan=5, sticky="w", padx=(8, 0), pady=(8, 0))
        tk.Label(option_row, text="입력 방식", fg="#52617A", bg="#FFFFFF").grid(row=3, column=0, sticky="w", pady=(8, 0))
        source_row = tk.Frame(option_row, bg="#FFFFFF")
        source_row.grid(row=3, column=1, columnspan=5, sticky="w", padx=(8, 0), pady=(8, 0))
        tk.Radiobutton(
            source_row,
            text="TXT로 분석 (LMK 불필요)",
            variable=self.force_pcf_var,
            value=False,
            bg="#FFFFFF",
            fg="#071B3D",
            activebackground="#FFFFFF",
            selectcolor="#FFFFFF",
            font=("Malgun Gothic", 10),
        ).pack(side="left", padx=(0, 18))
        tk.Radiobutton(
            source_row,
            text="PCF에서 TXT 추출 후 분석 (LMK 필요)",
            variable=self.force_pcf_var,
            value=True,
            bg="#FFFFFF",
            fg="#071B3D",
            activebackground="#FFFFFF",
            selectcolor="#FFFFFF",
            font=("Malgun Gothic", 10),
        ).pack(side="left")

        help_text = "LMK가 설치되지 않은 PC는 'TXT로 분석'을 선택하고 'TXT 파일 선택'으로 기존 *_Lxy.txt / *_WST.txt를 불러와 분석하세요. PCF가 자동 검색되지 않으면 'PCF 파일 선택'으로 직접 선택할 수 있습니다. 'PCF에서 TXT 추출 후 분석'은 PC별 LabSoft 설치/라이선스/자동제어 설정 영향을 받습니다. 생성 TXT는 PCF 이름 기준(예: 3M#1_Lxy.txt / 3M#1_WST.txt)으로 저장됩니다."
        tk.Label(input_card, text=help_text, fg="#536789", bg="#FFFFFF", wraplength=1040, justify="left").grid(
            row=4, column=0, columnspan=6, sticky="w", padx=(0, 0), pady=(8, 10)
        )

        action_row = tk.Frame(input_card, bg="#FFFFFF")
        action_row.grid(row=5, column=0, columnspan=6, sticky="ew", pady=(2, 0))
        action_row.grid_columnconfigure(6, weight=1)
        self.create_button = self.button(action_row, "분석", self.start, fill="#1769D1", fg="#FFFFFF", height=42)
        self.create_button.grid(row=0, column=0, sticky="w", padx=(0, 10), ipadx=10)
        self.stop_button = self.button(action_row, "중지", self.stop_work, fill="#FEE2E2", fg="#B91C1C", height=42)
        self.stop_button.grid(row=0, column=1, sticky="w", padx=(0, 10), ipadx=10)
        self.stop_button.configure(state="disabled")
        self.button(action_row, "Excel 열기", self.open_last_output, fill="#EEF5FF", fg="#1257B2").grid(row=0, column=2, sticky="w", padx=(0, 10), ipadx=8, ipady=4)
        self.button(action_row, "CSV 열기", self.open_last_csv, fill="#EEF5FF", fg="#1257B2").grid(row=0, column=3, sticky="w", padx=(0, 10), ipadx=8, ipady=4)
        self.button(action_row, "결과 폴더", self.open_output_folder, fill="#EEF5FF", fg="#1257B2").grid(row=0, column=4, sticky="w", padx=(0, 10), ipadx=8, ipady=4)
        self.button(action_row, "초기화", self.reset_form, fill="#F5F9FF", fg="#1257B2").grid(row=0, column=5, sticky="w", padx=(0, 16), ipadx=8, ipady=4)
        self.progress_bar = ttk.Progressbar(action_row, style="MPCD.Horizontal.TProgressbar", mode="determinate", maximum=100, value=0)
        self.progress_bar.grid(row=0, column=6, sticky="ew", padx=(0, 12))
        tk.Label(action_row, textvariable=self.status_var, fg="#334155", bg="#FFFFFF").grid(row=0, column=7, sticky="e")

        self.result_card = self.card(shell)
        self.result_card.pack(fill="both", expand=True, pady=(0, 10))
        self.result_title_label = tk.Label(
            self.result_card,
            text="15 x 35 MPCD 결과",
            font=("Malgun Gothic", 12, "bold"),
            fg="#071B3D",
            bg="#FFFFFF",
        )
        self.result_title_label.pack(anchor="w")
        self.result_meta_label = tk.Label(
            self.result_card,
            text=self.result_meta_text(),
            font=("Malgun Gothic", 10, "bold"),
            fg="#071B3D",
            bg="#FFFFFF",
            justify="left",
            wraplength=1040,
        )
        self.result_meta_label.pack(anchor="w", pady=(10, 0))
        self.result_detail_label = tk.Label(
            self.result_card,
            text="분석을 실행하면 Excel / CSV 결과를 여기서 바로 열 수 있습니다.",
            fg="#52617A",
            bg="#FFFFFF",
            wraplength=1040,
            justify="left",
        )
        self.result_detail_label.pack(anchor="w", pady=(4, 10))
        scale_row = tk.Frame(self.result_card, bg="#FFFFFF")
        scale_row.pack(fill="x", pady=(4, 12))
        tk.Label(scale_row, text="색상:", font=("Malgun Gothic", 10, "bold"), fg="#071B3D", bg="#FFFFFF").pack(side="left", padx=(0, 8))
        self.add_swatch(scale_row, "#F8696B", "-7")
        self.add_swatch(scale_row, "#FFFFFF", "3")
        self.add_swatch(scale_row, "#4D93D9", "13")
        preview_row = tk.Frame(self.result_card, bg="#FFFFFF")
        preview_row.pack(fill="x", pady=(0, 8))
        tk.Label(preview_row, text="결과 시료", font=("Malgun Gothic", 10, "bold"), fg="#071B3D", bg="#FFFFFF").pack(side="left", padx=(0, 8))
        self.result_sample_combo = ttk.Combobox(preview_row, textvariable=self.result_sample_var, values=(), width=28, state="readonly")
        self.result_sample_combo.pack(side="left")
        self.result_sample_combo.bind("<<ComboboxSelected>>", lambda _event: self.update_result_preview())
        self.result_table_frame = tk.Frame(self.result_card, bg="#FFFFFF")
        self.result_table_frame.pack(fill="both", expand=True, pady=(0, 10))
        self.result_table_frame.grid_rowconfigure(0, weight=1)
        self.result_table_frame.grid_columnconfigure(0, weight=1)
        self.result_text = tk.Text(
            self.result_table_frame,
            height=12,
            wrap="none",
            relief="solid",
            bd=1,
            bg="#FFFFFF",
            fg="#071B3D",
            font=("Consolas", 9),
        )
        self.result_text.grid(row=0, column=0, sticky="nsew")
        result_y_scroll = ttk.Scrollbar(self.result_table_frame, orient="vertical", command=self.result_text.yview)
        result_y_scroll.grid(row=0, column=1, sticky="ns")
        result_x_scroll = ttk.Scrollbar(self.result_table_frame, orient="horizontal", command=self.result_text.xview)
        result_x_scroll.grid(row=1, column=0, sticky="ew")
        self.result_text.configure(yscrollcommand=result_y_scroll.set, xscrollcommand=result_x_scroll.set, state="disabled")
        result_buttons = tk.Frame(self.result_card, bg="#FFFFFF")
        result_buttons.pack(fill="x")
        self.button(result_buttons, "Excel 결과 보기", self.open_last_output, fill="#1769D1", fg="#FFFFFF").pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.button(result_buttons, "CSV 결과 보기", self.open_last_csv, fill="#0F766E", fg="#FFFFFF").pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.button(result_buttons, "결과 폴더 열기", self.open_output_folder, fill="#EEF5FF", fg="#1257B2").pack(side="left", fill="x", expand=True)

    def maximize_window(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            width = self.winfo_screenwidth()
            height = self.winfo_screenheight()
            self.geometry(f"{width}x{height}+0+0")

    def card(self, parent: tk.Misc) -> tk.Frame:
        frame = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1)
        frame.configure(padx=16, pady=14)
        return frame

    def button(
        self,
        parent: tk.Misc,
        text: str,
        command: callable,
        fill: str = "#F8FAFC",
        fg: str = "#111827",
        height: int = 34,
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=fill,
            fg=fg,
            activebackground=fill,
            activeforeground=fg,
            disabledforeground=fg,
            relief="flat",
            bd=0,
            height=max(1, height // 20),
            cursor="hand2",
            font=("Malgun Gothic", 10, "bold"),
        )

    def add_picker_card(
        self,
        parent: tk.Misc,
        title: str,
        subtitle: str,
        value: str,
        command: callable,
        open_command: callable,
        grid_options: Optional[dict] = None,
    ) -> tk.Label:
        frame = self.card(parent)
        if grid_options:
            frame.grid(**grid_options)
        else:
            frame.pack(fill="x", pady=(0, 12))
        top = tk.Frame(frame, bg="#FFFFFF")
        top.pack(fill="x")
        tk.Label(top, text=title, font=("Malgun Gothic", 12, "bold"), fg="#111827", bg="#FFFFFF").pack(side="left")
        self.button(top, "↗ 보기", open_command, fill="#F1F5F9", fg="#334155").pack(side="right", padx=(6, 0))
        self.button(top, "＋ 선택", command, fill="#EEF2FF", fg="#1D4ED8").pack(side="right")
        tk.Label(frame, text=subtitle, fg="#64748B", bg="#FFFFFF").pack(anchor="w", pady=(5, 7))
        value_label = tk.Label(frame, text=value, fg="#0F172A", bg="#F8FAFC", anchor="w", padx=10, pady=7)
        value_label.pack(fill="x")
        return value_label

    def add_layout_button(self, parent: tk.Misc, label: str, value: str) -> None:
        button = self.button(parent, label, lambda v=value: self.set_layout(v), fill="#F8FAFC", fg="#111827")
        button.pack(side="left", fill="x", expand=True, padx=(0, 6))

    def add_swatch(self, parent: tk.Misc, color: str, label: str) -> None:
        item = tk.Frame(parent, bg="#FFFFFF")
        item.pack(side="left", padx=(0, 18))
        tk.Label(item, text="  ", bg=color, highlightbackground="#CBD5E1", highlightthickness=1).pack(side="left")
        tk.Label(item, text=f" {label}", fg="#475569", bg="#FFFFFF").pack(side="left")

    def display_path(self, path: Path) -> str:
        if not str(path):
            return "선택되지 않음"
        return path.name or str(path)

    def panel_roi_text(self) -> str:
        raw = self.panel_roi_var.get().strip()
        try:
            value = float(raw)
        except ValueError:
            return raw or "-"
        if value <= 1:
            value *= 100
        return f"{value:.1f}%"

    def result_meta_text(self) -> str:
        return (
            f"배열: 15 x 35 Signed MPCD    Panel ROI: {self.panel_roi_text()}    "
            "색상: -MPCD Red / +MPCD Blue"
        )

    def refresh_cards(self) -> None:
        if self.folder_value_label:
            self.folder_value_label.configure(text=self.display_path(Path(self.folder_var.get())))
        if self.reference_value_label:
            self.reference_value_label.configure(text=self.display_path(Path(self.reference_var.get())))
        if self.output_value_label:
            output = Path(self.output_var.get())
            self.output_value_label.configure(text=f"{output.name} / {output.with_suffix('.csv').name}")
        if self.result_meta_label:
            self.result_meta_label.configure(text=self.result_meta_text())
        if self.manual_pcf_label:
            if self.manual_pcf_files:
                names = ", ".join(path.name for path in self.manual_pcf_files[:4])
                if len(self.manual_pcf_files) > 4:
                    names += f" 외 {len(self.manual_pcf_files) - 4}개"
                self.manual_pcf_label.configure(text=f"수동 선택 PCF: {names}")
            else:
                self.manual_pcf_label.configure(text="수동 선택 PCF: 없음")

    def refresh_sample_options(self) -> None:
        if not self.sample_listbox:
            return
        folder = Path(self.folder_var.get())
        selected_ids = set(self.get_selected_sample_ids())
        pcf_files = discover_pcf_files(folder, self.manual_pcf_files)
        if pcf_files:
            values = ["전체"] + [f"Sample {sample_id} - {path.name}" for sample_id, path in pcf_files.items()]
        else:
            sample_ids = discover_sample_ids(folder)
            values = ["전체"] + [f"Sample {sample_id}" for sample_id in sample_ids]
        values = values[1:]
        self.sample_listbox.delete(0, tk.END)
        for value in values:
            self.sample_listbox.insert(tk.END, value)
        for index, value in enumerate(values):
            sample_id = parse_sample_selection(value)
            if sample_id in selected_ids:
                self.sample_listbox.selection_set(index)

    def get_selected_sample_ids(self) -> List[int]:
        if not self.sample_listbox:
            return []
        sample_ids: List[int] = []
        for index in self.sample_listbox.curselection():
            sample_id = parse_sample_selection(self.sample_listbox.get(index))
            if sample_id is not None:
                sample_ids.append(sample_id)
        return sample_ids

    def get_sample_selection(self) -> str:
        if not self.sample_listbox:
            return "전체"
        selected_values = [self.sample_listbox.get(index) for index in self.sample_listbox.curselection()]
        return "\n".join(selected_values) if selected_values else "전체"

    def select_all_samples(self) -> None:
        if self.sample_listbox:
            self.sample_listbox.selection_set(0, tk.END)
            self.status_var.set("시료 전체를 선택했습니다.")

    def clear_sample_selection(self) -> None:
        if self.sample_listbox:
            self.sample_listbox.selection_clear(0, tk.END)
            self.status_var.set("시료 선택을 해제했습니다. 선택 없음은 전체 분석입니다.")

