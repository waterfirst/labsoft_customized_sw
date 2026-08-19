
def export_pcf_inputs_with_labsoft(
    folder: Path,
    pcf_pairs: Sequence[Tuple[int, Path]],
    progress: Optional[callable] = None,
    labsoft_path: str = "",
    overwrite_existing: bool = False,
) -> None:
    if os.name != "nt":
        raise RuntimeError("LabSoft4 ActiveX export is only available on Windows.")
    if not pcf_pairs:
        return
    for stale_dir in folder.glob("mpcd_pcf_*"):
        if stale_dir.is_dir():
            shutil.rmtree(stale_dir, ignore_errors=True)
    with tempfile.TemporaryDirectory(prefix="mpcd_pcf_", dir=str(folder)) as dump_name:
        dump_root = Path(dump_name)
        for sample_id, pcf_path in pcf_pairs:
            sample_dir: Optional[Path] = None
            try:
                sizes = run_labsoft_dump([(sample_id, pcf_path)], dump_root, progress, labsoft_path)
                if sample_id not in sizes:
                    raise RuntimeError(f"LabSoft4 did not return export size for Sample {sample_id}.")
                rows, cols = sizes[sample_id]
                sample_dir = dump_root / str(sample_id)
                lxy_path, wst_path = preferred_raw_txt_paths(folder, sample_id, pcf_path)
                crop_bounds = None
                if not pcf_looks_cropped(pcf_path):
                    crop_bounds = detect_luminance_crop_bounds(sample_dir / "lxy_1.bin", rows, cols)
                    if crop_bounds is not None and progress:
                        r1, r2, c1, c2 = crop_bounds
                        progress(
                            f"Sample {sample_id}: {pcf_path.name} auto crop "
                            f"{r2 - r1 + 1} x {c2 - c1 + 1}"
                        )
                    elif progress:
                        progress(f"Sample {sample_id}: auto crop area not found, use full PCF")
                metadata_path = metadata_path_for_raw(lxy_path)
                metadata = {
                    "sample_id": sample_id,
                    "pcf_name": pcf_path.name,
                    "pcf_path": str(pcf_path),
                    "lxy_name": lxy_path.name,
                    "wst_name": wst_path.name,
                    "pcf_looks_cropped": pcf_looks_cropped(pcf_path),
                    "auto_crop_bounds": list(crop_bounds) if crop_bounds is not None else None,
                    "rotate_final_matrix_180": False,
                }
                if overwrite_existing or not lxy_path.exists():
                    if progress:
                        progress(f"Sample {sample_id}: create {lxy_path.name}")
                    write_labsoft_txt_from_dumps(
                        [sample_dir / f"lxy_{index}.bin" for index in (1, 2, 3)],
                        rows,
                        cols,
                        lxy_path,
                        crop_bounds,
                    )
                if overwrite_existing or not metadata_path.exists():
                    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
                if overwrite_existing or not wst_path.exists():
                    if progress:
                        progress(f"Sample {sample_id}: create {wst_path.name}")
                    write_labsoft_txt_from_dumps(
                        [sample_dir / f"wst_{index}.bin" for index in (1, 2, 3)],
                        rows,
                        cols,
                        wst_path,
                        crop_bounds,
                    )
            except OSError as exc:
                if exc.errno == errno.ENOSPC:
                    raise OSError(
                        errno.ENOSPC,
                        f"디스크 여유 공간이 부족합니다. Data 폴더의 불필요한 파일 또는 mpcd_pcf_* 임시 폴더를 삭제한 뒤 다시 실행해 주세요. folder={folder}",
                    ) from exc
                raise
            finally:
                if sample_dir is not None and sample_dir.exists():
                    shutil.rmtree(sample_dir, ignore_errors=True)


def ensure_raw_txt_inputs(
    folder: Path,
    sample_selection: str,
    progress: Optional[callable] = None,
    labsoft_path: str = "",
    force_pcf: bool = False,
    manual_pcf_files: Optional[Sequence[Path]] = None,
) -> None:
    selected_ids = parse_sample_selection_ids(sample_selection)
    pcf_files = discover_pcf_files(folder, manual_pcf_files)
    sample_ids = selected_ids if selected_ids is not None else sorted(pcf_files)
    missing_pairs: List[Tuple[int, Path]] = []
    missing_without_pcf: List[int] = []
    for sample_id in sample_ids:
        pcf_path = pcf_files.get(sample_id)
        lxy_path, wst_path = existing_raw_txt_paths(folder, sample_id, pcf_path)
        if force_pcf and pcf_path:
            missing_pairs.append((sample_id, pcf_path))
            continue
        if lxy_path.exists() and wst_path.exists():
            continue
        if force_pcf and pcf_path:
            missing_pairs.append((sample_id, pcf_path))
        else:
            missing_without_pcf.append(sample_id)
    if missing_pairs:
        export_pcf_inputs_with_labsoft(folder, missing_pairs, progress, labsoft_path, force_pcf)
    if missing_without_pcf:
        sample_text = ", ".join(f"Sample {sample_id}" for sample_id in missing_without_pcf)
        raise FileNotFoundError(
            f"{sample_text}: *_Lxy.txt / *_WST.txt 파일을 찾지 못했습니다. "
            "PCF 자동 추출이 필요한 경우 'PCF에서 새로 추출'을 체크하고, "
            "LabSoft가 이 PC에서 수동으로 정상 실행/PCF 열기 가능한지 먼저 확인해 주세요."
        )


def selected_sample_ids_for_folder(folder: Path, sample_selection: str, manual_pcf_files: Optional[Sequence[Path]] = None) -> List[int]:
    selected_ids = parse_sample_selection_ids(sample_selection)
    if selected_ids is not None:
        return selected_ids
    return discover_sample_ids(folder, manual_pcf_files)


def existing_txt_sample_ids(folder: Path, sample_selection: str, manual_pcf_files: Optional[Sequence[Path]] = None) -> List[int]:
    pcf_files = discover_pcf_files(folder, manual_pcf_files)
    return [
        sample_id
        for sample_id in selected_sample_ids_for_folder(folder, sample_selection, manual_pcf_files)
        if all(path.exists() for path in existing_raw_txt_paths(folder, sample_id, pcf_files.get(sample_id)))
    ]


def col_to_num(col: str) -> int:
    value = 0
    for ch in col:
        value = value * 26 + ord(ch.upper()) - 64
    return value


def col_letter(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def cell_ref(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def split_cell_ref(ref: str) -> Tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        raise ValueError(f"잘못된 cell ref: {ref}")
    return int(match.group(2)), col_to_num(match.group(1))


def parse_number(text: str) -> float:
    return float(text.strip())


def split_numeric_parts(line: str) -> Optional[List[str]]:
    line = line.strip().lstrip("\ufeff")
    if not line:
        return None
    parts = [part for part in re.split(r"[\s,]+", line) if part]
    if not parts:
        return None
    try:
        parse_number(parts[0])
    except ValueError:
        return None
    return parts


def numeric_rows(path: Path) -> Iterable[Tuple[int, List[str]]]:
    with path.open("r", encoding="utf-8-sig", errors="ignore") as handle:
        for line_no, line in enumerate(handle, start=1):
            parts = split_numeric_parts(line)
            if parts is not None:
                yield line_no, parts


def detect_raw_geometry(lxy_path: Path, progress: Optional[callable] = None) -> RawGeometry:
    raw_width = -1
    width = -1
    height = 0
    sample_step = 1
    count = 0
    sum_x = 0.0
    sum_y = 0.0

    if progress:
        progress(f"{lxy_path.name}: 패널 영역 확인 중...")
    for _, parts in numeric_rows(lxy_path):
        if raw_width < 0:
            raw_width = len(parts)
            if raw_width % 3 != 0:
                raise ValueError(f"{lxy_path.name}: L, Wx, Wy 3채널 raw 형식이 아닙니다.")
            width = raw_width // 3
            sample_step = max(1, int(round(width / 1000)))
        elif len(parts) != raw_width:
            raise ValueError(f"{lxy_path.name}: raw 행의 열 개수가 서로 다릅니다.")

        if height % sample_step == 0:
            for x in range(0, width, sample_step):
                try:
                    lum = parse_number(parts[x * 3])
                except ValueError:
                    continue
                if lum >= LUMINANCE_THRESHOLD:
                    count += 1
                    sum_x += x
                    sum_y += height
        height += 1

    if raw_width <= 0 or height <= 0:
        raise ValueError(f"{lxy_path.name}: 숫자 raw matrix를 찾지 못했습니다.")
    if count == 0:
        raise ValueError(f"{lxy_path.name}: {LUMINANCE_THRESHOLD:g} nit 이상 패널 영역을 찾지 못했습니다.")

    cx = sum_x / count
    cy = sum_y / count
    xx = yy = xy = 0.0
    row = 0
    for _, parts in numeric_rows(lxy_path):
        if row % sample_step == 0:
            dy = row - cy
            for x in range(0, width, sample_step):
                try:
                    lum = parse_number(parts[x * 3])
                except ValueError:
                    continue
                if lum >= LUMINANCE_THRESHOLD:
                    dx = x - cx
                    xx += dx * dx
                    yy += dy * dy
                    xy += dx * dy
        row += 1

    theta = 0.5 * math.atan2(2 * xy, xx - yy)
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)
    u_min = math.inf
    u_max = -math.inf
    v_min = math.inf
    v_max = -math.inf
    row = 0
    for _, parts in numeric_rows(lxy_path):
        if row % sample_step == 0:
            for x in range(0, width, sample_step):
                try:
                    lum = parse_number(parts[x * 3])
                except ValueError:
                    continue
                if lum >= LUMINANCE_THRESHOLD:
                    dx = x - cx
                    dy = row - cy
                    u = cos_t * dx + sin_t * dy
                    v = -sin_t * dx + cos_t * dy
                    u_min = min(u_min, u)
                    u_max = max(u_max, u)
                    v_min = min(v_min, v)
                    v_max = max(v_max, v)
        row += 1

    return RawGeometry(width, height, sample_step, cx, cy, cos_t, sin_t, u_min, u_max, v_min, v_max, theta * 180.0 / math.pi)


def apply_panel_roi(geometry: RawGeometry, panel_roi: float) -> RawGeometry:
    roi = max(0.01, min(1.0, float(panel_roi)))
    u_span = geometry.u_max - geometry.u_min
    v_span = geometry.v_max - geometry.v_min
    u_inset = u_span * (1.0 - roi) * 0.5
    v_inset = v_span * (1.0 - roi) * 0.5
    return RawGeometry(
        geometry.width,
        geometry.height,
        geometry.sample_step,
        geometry.cx,
        geometry.cy,
        geometry.cos,
        geometry.sin,
        geometry.u_min + u_inset,
        geometry.u_max - u_inset,
        geometry.v_min + v_inset,
        geometry.v_max - v_inset,
        geometry.tilt_deg,
    )


def xy_to_uv1960(x: float, y: float) -> Tuple[float, float]:
    denom = 3.0 + 12.0 * y - 2.0 * x
    if denom == 0:
        return 0.0, 0.0
    return 4.0 * x / denom, 6.0 * y / denom


def color_to_uv1960(x: float, y: float, mode: str) -> Tuple[float, float]:
    if mode == "uv1976":
        return x, y / 1.5
    return xy_to_uv1960(x, y)


def signed_duv(measured_u: float, measured_v: float, ref_u: float, ref_v: float) -> float:
    distance = math.hypot(measured_u - ref_u, measured_v - ref_v)
    sign = 1.0 if (ref_v < measured_v or ref_u > measured_u) else -1.0
    return distance * sign


def planckian_xy_from_cct(t: float) -> Tuple[float, float]:
    if t <= 0:
        return 0.0, 0.0
    if t < 4000:
        x = -0.2661239e9 / (t**3) - 0.234358e6 / (t**2) + 0.8776956e3 / t + 0.17991
    else:
        x = -3.0258469e9 / (t**3) + 2.1070379e6 / (t**2) + 0.2226347e3 / t + 0.24039
    if t < 2222:
        y = -1.1063814 * (x**3) - 1.3481102 * (x**2) + 2.18555832 * x - 0.20219683
    elif t < 4000:
        y = -0.9549476 * (x**3) - 1.37418593 * (x**2) + 2.09137015 * x - 0.16748867
    elif t < 25000:
        y = 3.081758 * (x**3) - 5.8733867 * (x**2) + 3.75112997 * x - 0.37001483
    else:
        y = 0.0
    return x, y


def detect_color_mode(sums: Sequence[Sequence[dict]]) -> str:
    total_x = total_y = 0.0
    total_pixels = 0
    for row in sums:
        for cell in row:
            total_x += cell["wx"]
            total_y += cell["wy"]
            total_pixels += cell["pixels"]
    if total_pixels <= 0:
        return "xy"
    avg_x = total_x / total_pixels
    avg_y = total_y / total_pixels
    if 0.15 <= avg_x <= 0.25 and 0.35 <= avg_y <= 0.55:
        return "uv1976"
    return "xy"


def raw_matrix_summary(matrix: Sequence[Sequence[Optional[float]]]) -> List[Tuple[str, Optional[float]]]:
    def avg(row_start: int, row_end: int, columns: Iterable[int]) -> Optional[float]:
        values = []
        for row in range(row_start, row_end + 1):
            if row < 1 or row > len(matrix):
                continue
            line = matrix[row - 1]
            for column in columns:
                if column < 1 or column > len(line):
                    continue
                value = line[column - 1]
                if value is not None and math.isfinite(value):
                    values.append(value)
        return statistics.mean(values) if values else None

    top_edge = avg(6, 10, [1, 15])
    top_center = avg(6, 10, range(6, 11))
    mid_edge = avg(16, 20, [1, 15])
    mid_center = avg(16, 20, range(6, 11))
    bottom_edge = avg(26, 30, [1, 15])
    bottom_center = avg(26, 30, range(6, 11))

    def delta(edge: Optional[float], center: Optional[float]) -> Optional[float]:
        return center - edge if edge is not None and center is not None else None

    deltas = [delta(top_edge, top_center), delta(mid_edge, mid_center), delta(bottom_edge, bottom_center)]
    ave = statistics.mean([value for value in deltas if value is not None]) if any(value is not None for value in deltas) else None
    return [
        ("상EDGE", top_edge),
        ("상CENTER", top_center),
        ("중EDGE", mid_edge),
        ("중CENTER", mid_center),
        ("하EDGE", bottom_edge),
        ("하CENTER", bottom_center),
        ("상delta", deltas[0]),
        ("중delta", deltas[1]),
        ("하delta", deltas[2]),
        ("AVE", ave),
    ]


def calculate_raw_matrix(
    lxy_path: Path,
    wst_path: Path,
    panel_roi: float,
    progress: Optional[callable] = None,
) -> Tuple[List[List[Optional[float]]], List[Tuple[str, Optional[float]]]]:
    geometry = detect_raw_geometry(lxy_path, progress)
    grid_geometry = apply_panel_roi(geometry, panel_roi)
    u_span = grid_geometry.u_max - grid_geometry.u_min
    v_span = grid_geometry.v_max - grid_geometry.v_min
    u_cell = u_span / RAW_GRID_COLS
    v_cell = v_span / RAW_GRID_ROWS
    sums = [[{"l": 0.0, "wx": 0.0, "wy": 0.0, "t": 0.0, "pixels": 0} for _ in range(RAW_GRID_COLS)] for _ in range(RAW_GRID_ROWS)]
    wst_width = -1
    wst_stride = 1
    wst_offset = 0
    row_index = 0

    if progress:
        progress(f"{lxy_path.name}: MPCD 계산 중...")
    for lxy_item, wst_item in zip_longest(numeric_rows(lxy_path), numeric_rows(wst_path)):
        if lxy_item is None or wst_item is None:
            raise ValueError("LXY raw와 WST raw의 숫자 행 개수가 서로 다릅니다.")
        lxy_line_no, lxy_parts = lxy_item
        wst_line_no, wst_parts = wst_item
        if len(lxy_parts) != geometry.width * 3:
            raise ValueError(f"{lxy_path.name}: {lxy_line_no}행 열 개수가 LXY 3채널 형식과 다릅니다.")
        if wst_width < 0:
            wst_width = len(wst_parts)
            if wst_width == geometry.width:
                wst_stride = 1
                wst_offset = 0
            elif wst_width == geometry.width * 3:
                wst_stride = 3
                wst_offset = 2
            else:
                raise ValueError(f"{wst_path.name}: WST 열 개수는 T 단일 채널 또는 W/S/T 3채널이어야 합니다.")
        elif len(wst_parts) != wst_width:
            raise ValueError(f"{wst_path.name}: {wst_line_no}행 열 개수가 첫 데이터 행과 다릅니다.")

        for x in range(geometry.width):
            base = x * 3
            try:
                lum = parse_number(lxy_parts[base])
            except ValueError:
                continue
            if lum < LUMINANCE_THRESHOLD:
                continue
            dx = x - geometry.cx
            dy = row_index - geometry.cy
            u_pos = geometry.cos * dx + geometry.sin * dy
            v_pos = -geometry.sin * dx + geometry.cos * dy
            if u_pos < grid_geometry.u_min or u_pos > grid_geometry.u_max or v_pos < grid_geometry.v_min or v_pos > grid_geometry.v_max:
                continue
            col = int(math.floor((u_pos - grid_geometry.u_min) / u_cell))
            row = int(math.floor((v_pos - grid_geometry.v_min) / v_cell))
            if row < 0 or row >= RAW_GRID_ROWS or col < 0 or col >= RAW_GRID_COLS:
                continue
            try:
                wx = parse_number(lxy_parts[base + 1])
                wy = parse_number(lxy_parts[base + 2])
                temp = parse_number(wst_parts[x * wst_stride + wst_offset])
            except (ValueError, IndexError):
                continue
            if not (math.isfinite(wx) and math.isfinite(wy) and math.isfinite(temp) and temp > 0):
                continue
            item = sums[row][col]
            item["l"] += lum
            item["wx"] += wx
            item["wy"] += wy
            item["t"] += temp
            item["pixels"] += 1
        row_index += 1

    color_mode = detect_color_mode(sums)
    grid = [[RawCell() for _ in range(RAW_GRID_COLS)] for _ in range(RAW_GRID_ROWS)]
    for row in range(RAW_GRID_ROWS):
        for col in range(RAW_GRID_COLS):
            item = sums[row][col]
            pixels = int(item["pixels"])
            if pixels == 0:
                continue
            wx = item["wx"] / pixels
            wy = item["wy"] / pixels
            temp = item["t"] / pixels
            measured_u, measured_v = color_to_uv1960(wx, wy, color_mode)
            ref_x, ref_y = planckian_xy_from_cct(temp)
            ref_u, ref_v = xy_to_uv1960(ref_x, ref_y)
            grid[row][col] = RawCell(signed_duv(measured_u, measured_v, ref_u, ref_v) / MPCD_UNIT)

    matrix: List[List[Optional[float]]] = []
    for photo_row in range(RAW_GRID_COLS - 1, -1, -1):
        line = []
        for photo_col in range(RAW_GRID_ROWS):
            line.append(grid[photo_col][photo_row].mpcd)
        matrix.append(line)
    return matrix, raw_matrix_summary(matrix)


def parse_sample_selection(sample_selection: str) -> Optional[int]:
    sample_ids = parse_sample_selection_ids(sample_selection)
    return sample_ids[0] if sample_ids else None


def parse_sample_selection_ids(sample_selection: str) -> Optional[List[int]]:
    if not sample_selection or sample_selection.strip() == "전체":
        return None
    sample_ids: List[int] = []
    for part in re.split(r"[\n,;]+", sample_selection):
        text = part.strip()
        if not text:
            continue
        if text == "전체":
            return None
        match = re.search(r"Sample\s+(\d+)", text, re.IGNORECASE)
        if not match:
            match = re.search(r"\b(\d+)\b", text)
        if match:
            sample_id = int(match.group(1))
            if sample_id not in sample_ids:
                sample_ids.append(sample_id)
    return sample_ids or None


def read_raw_results(
    folder: Path,
    panel_roi: float,
    progress: Optional[callable] = None,
    sample_selection: str = "전체",
    manual_pcf_files: Optional[Sequence[Path]] = None,
) -> List[SampleResult]:
    labsoft_setting = os.environ.get("MPCD_LABSOFT_PATH", "")
    force_pcf = os.environ.get("MPCD_FORCE_PCF", "0") == "1"
    ensure_raw_txt_inputs(folder, sample_selection, progress, labsoft_setting, force_pcf, manual_pcf_files)
    sample_ids = discover_sample_ids(folder, manual_pcf_files)
    pcf_files = discover_pcf_files(folder, manual_pcf_files)
    selected_ids = parse_sample_selection_ids(sample_selection)
    if selected_ids is not None:
        selected_set = set(selected_ids)
        sample_ids = [sample_id for sample_id in sample_ids if sample_id in selected_set]
    sample_ids = [
        sample_id
        for sample_id in sample_ids
        if all(path.exists() for path in existing_raw_txt_paths(folder, sample_id, pcf_files.get(sample_id)))
    ]
    if not sample_ids:
        raise ValueError(f"{folder}에서 *_Lxy.txt 파일을 찾지 못했습니다.")
    results: List[SampleResult] = []
    pcf_ids = set(pcf_files)
    for sample_id in sample_ids:
        lxy_file, wst_file = existing_raw_txt_paths(folder, sample_id, pcf_files.get(sample_id))
        if not wst_file.exists():
            raise FileNotFoundError(f"WST 파일을 찾을 수 없습니다: {wst_file}")
        if progress:
            progress(f"Sample {sample_id}: raw MPCD 계산 중...")
        matrix, summary = calculate_raw_matrix(lxy_file, wst_file, panel_roi, progress)
        matrix = apply_hole_mask(matrix)
        summary = raw_matrix_summary(matrix)
        source_text = "PCF export" if force_pcf and sample_id in pcf_ids else "raw txt"
        display_name = display_name_from_path(pcf_files[sample_id]) if sample_id in pcf_files else display_name_from_path(lxy_file).replace(" Lxy", "")
        results.append(
            SampleResult(
                sample_id=sample_id,
                values=matrix,
                side_summary=summary,
                source=f"{source_text} / ROI {panel_roi:g}",
                display_name=display_name,
                lxy_file=lxy_file,
                wst_file=wst_file,
            )
        )
    return results


def parse_shared_strings(zip_file: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings = []
    for si in root.findall("m:si", ns):
        texts = [node.text or "" for node in si.findall(".//m:t", ns)]
        strings.append("".join(texts))
    return strings


def read_sheet_cells(xlsx_path: Path, sheet_name: str = "sheet1") -> Dict[Tuple[int, int], object]:
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    cells: Dict[Tuple[int, int], object] = {}
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = parse_shared_strings(z)
        sheet_path = f"xl/worksheets/{sheet_name}.xml"
        if sheet_path not in z.namelist():
            raise ValueError(f"{xlsx_path.name}에서 {sheet_path}를 찾지 못했습니다.")
        for _, elem in ET.iterparse(z.open(sheet_path), events=("end",)):
            if not elem.tag.endswith("c"):
                continue
            ref = elem.attrib.get("r", "")
            if not ref:
                elem.clear()
                continue
            row, col = split_cell_ref(ref)
            value_elem = elem.find(ns_main + "v")
            if value_elem is None or value_elem.text is None:
                elem.clear()
                continue
            raw = value_elem.text
            if elem.attrib.get("t") == "s":
                value = shared_strings[int(raw)]
            else:
                try:
                    value = float(raw)
                except ValueError:
                    value = raw
            cells[(row, col)] = value
            elem.clear()
    return cells


def reference_position(sample_id: int) -> Tuple[int, int, int]:
    if sample_id <= 6:
        pos = sample_id - 1
        return REFERENCE_FIRST_TOP_ROW + pos * REFERENCE_BLOCK_STEP, REFERENCE_LEFT_LABEL_COL, REFERENCE_LEFT_GRID_COL
    pos = sample_id - 7
    return REFERENCE_FIRST_TOP_ROW + pos * REFERENCE_BLOCK_STEP, REFERENCE_RIGHT_LABEL_COL, REFERENCE_RIGHT_GRID_COL


def discover_reference_blocks(cells: Dict[Tuple[int, int], object]) -> Dict[int, Tuple[int, int, int]]:
    blocks: Dict[int, Tuple[int, int, int]] = {}
    if not cells:
        return blocks

    max_row = max(row for row, _ in cells)
    candidates = [
        (REFERENCE_LEFT_LABEL_COL, REFERENCE_LEFT_GRID_COL),
        (REFERENCE_RIGHT_LABEL_COL, REFERENCE_RIGHT_GRID_COL),
    ]
    for label_col, grid_col in candidates:
        for row in range(1, max_row + 1):
            label_value = cells.get((row, label_col))
            if not isinstance(label_value, (int, float)):
                continue
            sample_id = int(label_value)
            if abs(label_value - sample_id) > 1e-9:
                continue
            value_count = 0
            for r in range(REFERENCE_GRID_ROWS):
                for c in range(REFERENCE_GRID_COLS):
                    if isinstance(cells.get((row + r, grid_col + c)), (int, float)):
                        value_count += 1
            if value_count >= REFERENCE_GRID_COLS:
                blocks[sample_id] = (row, label_col, grid_col)
    return dict(sorted(blocks.items()))


def read_reference_excel(folder: Path, reference_path: Optional[Path] = None) -> List[SampleResult]:
    reference_path = resolve_reference_excel(folder, reference_path)

    cells = read_sheet_cells(reference_path, "sheet1")
    reference_blocks = discover_reference_blocks(cells)
    folder_sample_ids = discover_sample_ids(folder)
    pcf_files = discover_pcf_files(folder)
    sample_ids = folder_sample_ids or sorted(reference_blocks)
    if not sample_ids:
        raise ValueError("기준 Excel에서 시료 블록을 찾지 못했습니다.")

    results: List[SampleResult] = []

    for sample_id in sample_ids:
        if sample_id not in reference_blocks:
            raise ValueError(
                f"Sample {sample_id}의 기준 MPCD 블록을 {reference_path.name}에서 찾지 못했습니다. "
                "시료 수가 늘어난 경우 기준 Excel에도 해당 시료 결과가 있어야 합니다."
            )
        top_row, label_col, grid_col = reference_blocks[sample_id]
        title_value = cells.get((top_row - 1, label_col + 1)) or cells.get((top_row - 1, grid_col))
        display_name = str(title_value) if isinstance(title_value, str) else f"Sample {sample_id}"
        values: List[List[Optional[float]]] = []
        for r in range(REFERENCE_GRID_ROWS):
            row_values: List[Optional[float]] = []
            for c in range(REFERENCE_GRID_COLS):
                value = cells.get((top_row + r, grid_col + c))
                row_values.append(float(value) if isinstance(value, (int, float)) else None)
            values.append(row_values)

        side_summary: List[Tuple[str, Optional[float]]] = []
        for index, label in enumerate(SUMMARY_LABELS):
            row = top_row + 25 + index
            summary_label = cells.get((row, label_col + SUMMARY_LABEL_OFFSET_COL))
            summary_value = cells.get((row, label_col + SUMMARY_VALUE_OFFSET_COL))
            if isinstance(summary_label, str):
                label = summary_label
            side_summary.append((label, float(summary_value) if isinstance(summary_value, (int, float)) else None))

        results.append(
            SampleResult(
                sample_id=sample_id,
                values=values,
                side_summary=side_summary,
                source=reference_path.name,
                display_name=display_name,
                lxy_file=existing_raw_txt_paths(folder, sample_id, pcf_files.get(sample_id))[0],
                wst_file=existing_raw_txt_paths(folder, sample_id, pcf_files.get(sample_id))[1],
            )
        )

    return results


def flat_values(values: Sequence[Sequence[Optional[float]]]) -> List[float]:
    return [v for row in values for v in row if v is not None and math.isfinite(v)]


def summarize_values(values: Sequence[Sequence[Optional[float]]]) -> Dict[str, Optional[float]]:
    data = flat_values(values)
    if not data:
        return {"count": 0, "mean": None, "std": None, "min": None, "max": None}
    return {
        "count": len(data),
        "mean": statistics.mean(data),
        "std": statistics.stdev(data) if len(data) > 1 else 0.0,
        "min": min(data),
        "max": max(data),
    }


def compute_side_summary(values: Sequence[Sequence[Optional[float]]]) -> List[Tuple[str, Optional[float]]]:
    def mean_cells(rows: Iterable[int], cols: Iterable[int]) -> Optional[float]:
        data = []
        row_list = list(rows)
        col_list = list(cols)
        for r in row_list:
            if r < 0 or r >= len(values):
                continue
            for c in col_list:
                if c < 0 or c >= len(values[r]):
                    continue
                value = values[r][c]
                if value is not None and math.isfinite(value):
                    data.append(value)
        return statistics.mean(data) if data else None

    rows = len(values)
    cols = len(values[0]) if rows else 0
    top = range(0, max(1, rows // 3))
    mid = range(max(0, rows // 3), max(1, rows * 2 // 3))
    bottom = range(max(0, rows * 2 // 3), rows)
    edge_cols = list(range(0, min(3, cols))) + list(range(max(0, cols - 3), cols))
    center_cols = range(max(0, cols // 2 - 2), min(cols, cols // 2 + 3))

    top_edge = mean_cells(top, edge_cols)
    top_center = mean_cells(top, center_cols)
    mid_edge = mean_cells(mid, edge_cols)
    mid_center = mean_cells(mid, center_cols)
    bottom_edge = mean_cells(bottom, edge_cols)
    bottom_center = mean_cells(bottom, center_cols)

    def delta(edge: Optional[float], center: Optional[float]) -> Optional[float]:
        return center - edge if edge is not None and center is not None else None

    all_mean = statistics.mean(flat_values(values)) if flat_values(values) else None
    return [
        ("상EDGE", top_edge),
        ("상CENTER", top_center),
        ("중EDGE", mid_edge),
        ("중CENTER", mid_center),
        ("하EDGE", bottom_edge),
        ("하CENTER", bottom_center),
        ("상delta", delta(top_edge, top_center)),
        ("중delta", delta(mid_edge, mid_center)),
        ("하delta", delta(bottom_edge, bottom_center)),
        ("AVE", all_mean),
    ]


def mix_hex(left: str, right: str, ratio: float) -> str:
    ratio = max(0.0, min(1.0, ratio))
    parts = []
    for i in range(0, 6, 2):
        a = int(left[i : i + 2], 16)
        b = int(right[i : i + 2], 16)
        parts.append(f"{round(a + (b - a) * ratio):02X}")
    return "".join(parts)


def interpolate_color(value: Optional[float]) -> str:
    if value is None or not math.isfinite(value):
        return "D9D9D9"
    if value <= COLOR_MID_VALUE:
        clipped = max(COLOR_MIN_VALUE, min(COLOR_MID_VALUE, value))
        ratio = (clipped - COLOR_MIN_VALUE) / (COLOR_MID_VALUE - COLOR_MIN_VALUE)
        return mix_hex(COLOR_MIN, COLOR_MID, ratio)
    clipped = max(COLOR_MID_VALUE, min(COLOR_MAX_VALUE, value))
    ratio = (clipped - COLOR_MID_VALUE) / (COLOR_MAX_VALUE - COLOR_MID_VALUE)
    return mix_hex(COLOR_MID, COLOR_MAX, ratio)


def xml_text(value: object) -> str:
    return escape(str(value))


class WorkbookBuilder:
    def __init__(self) -> None:
        self.color_to_style: Dict[str, int] = {}
        self.colors: List[str] = []

    def style_for_color(self, color: str) -> int:
        if color not in self.color_to_style:
            self.color_to_style[color] = 4 + len(self.colors)
            self.colors.append(color)
        return self.color_to_style[color]

    def cell(self, row: int, col: int, value: object, style: Optional[int] = None) -> str:
        style_attr = f' s="{style}"' if style is not None else ""
        ref = cell_ref(row, col)
        if value is None:
            return f'<c r="{ref}"{style_attr}/>'
        if isinstance(value, (int, float)) and math.isfinite(float(value)):
            return f'<c r="{ref}"{style_attr}><v>{float(value):.10g}</v></c>'
        return f'<c r="{ref}"{style_attr} t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'

    @staticmethod
    def row(row_index: int, cells: Sequence[str]) -> str:
        return f'<row r="{row_index}">{"".join(cells)}</row>' if cells else ""

    def placements(self, results: Sequence[SampleResult], layout: str, time_groups: Sequence[str]) -> List[Placement]:
        placements: List[Placement] = []
        if layout == "excel_style":
            for index, sample in enumerate(results):
                page = index // 12
                within_page = index % 12
                side = within_page // 6
                pos = within_page % 6
                label_col = REFERENCE_LEFT_LABEL_COL if side == 0 else REFERENCE_RIGHT_LABEL_COL
                grid_col = REFERENCE_LEFT_GRID_COL if side == 0 else REFERENCE_RIGHT_GRID_COL
                page_top = REFERENCE_FIRST_TOP_ROW + page * (6 * REFERENCE_BLOCK_STEP + 4)
                placements.append(Placement(sample, page_top + pos * REFERENCE_BLOCK_STEP, label_col, grid_col))
            return placements

        if layout == "vertical":
            top = 2
            for sample in results:
                placements.append(Placement(sample, top, 1, 2))
                top += sample.rows + 5
            return placements

        groups = list(time_groups) if time_groups else ["평가시간 1"]
        for index, sample in enumerate(results):
            group_index = min(len(groups) - 1, index * len(groups) // len(results))
            in_group_index = index - math.floor(group_index * len(results) / len(groups))
            row_pos = in_group_index
            base_col = 1 + group_index * 19
            placements.append(
                Placement(
                    sample=sample,
                    top_row=3 + row_pos * (sample.rows + 5),
                    label_col=base_col,
                    grid_col=base_col + 1,
                    group_label=groups[group_index],
                )
            )
        return placements

    def make_map_sheet(self, results: Sequence[SampleResult], layout: str, time_groups: Sequence[str]) -> str:
        rows: List[str] = []
        merges: List[str] = []
        placements = self.placements(results, layout, time_groups)

        group_headers = {}
        for placement in placements:
            if placement.group_label and placement.group_label not in group_headers:
                group_headers[placement.group_label] = placement.label_col
        for label, col in group_headers.items():
            merges.append(f'<mergeCell ref="{cell_ref(1, col)}:{cell_ref(1, col + 17)}"/>')
            rows.append(self.row(1, [self.cell(1, col, label, 1)]))

        for placement in placements:
            sample = placement.sample
            title_row = placement.top_row - 1
            title_end_col = placement.grid_col + sample.cols - 1
            merges.append(f'<mergeCell ref="{cell_ref(title_row, placement.label_col)}:{cell_ref(title_row, title_end_col)}"/>')
            rows.append(self.row(title_row, [self.cell(title_row, placement.label_col, sample_title(sample), 1)]))

            for r, source_row in enumerate(sample.values):
                cells = []
                if r == 0:
                    cells.append(self.cell(placement.top_row, placement.label_col, sample.sample_id, 2))
                for c, value in enumerate(source_row):
                    cells.append(self.cell(placement.top_row + r, placement.grid_col + c, value, self.style_for_color(interpolate_color(value))))
                rows.append(self.row(placement.top_row + r, cells))

            summary = sample.side_summary or compute_side_summary(sample.values)
            for idx, (label, value) in enumerate(summary):
                row_index = placement.top_row + 25 + idx
                rows.append(
                    self.row(
                        row_index,
                        [
                            self.cell(row_index, placement.grid_col + sample.cols + 1, label, 3),
                            self.cell(row_index, placement.grid_col + sample.cols + 2, value, 3),
                        ],
                    )
                )

        max_row = max((p.top_row + p.sample.rows + 12 for p in placements), default=1)
        max_col = max((p.grid_col + p.sample.cols + 3 for p in placements), default=1)
        cols = []
        for col in range(1, max_col + 1):
            width = 5 if col % 19 not in (1, 17, 18, 0) else 10
            cols.append(f'<col min="{col}" max="{col}" width="{width}" customWidth="1"/>')

        merge_xml = f'<mergeCells count="{len(merges)}">{"".join(merges)}</mergeCells>' if merges else ""
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="A1:{cell_ref(max_row, max_col)}"/>'
            '<sheetViews><sheetView workbookViewId="0" zoomScale="70" zoomScaleNormal="70"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16"/>'
            f'<cols>{"".join(cols)}</cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            f'{merge_xml}'
            '</worksheet>'
        )

    def make_new_result_sheet(self, results: Sequence[SampleResult]) -> str:
        row_cells: Dict[int, List[str]] = {}
        block_height = 65
        block_gap = 3
        summary_labels = [
            "상EDGE",
            "상CENTER",
            "중EDGE",
            "중CENTER",
            "하EDGE",
            "하CENTER",
            "상delta",
            "중delta",
            "하delta",
            "AVE",
        ]

        def band_means(values: Sequence[Sequence[Optional[float]]], start_row: int, end_row: int) -> List[Optional[float]]:
            band: List[Optional[float]] = []
            cols = len(values[0]) if values else 0
            for col in range(cols):
                data = []
                for row in range(start_row - 1, end_row):
                    if 0 <= row < len(values) and col < len(values[row]):
                        value = values[row][col]
                        if value is not None and math.isfinite(value):
                            data.append(value)
                band.append(statistics.mean(data) if data else None)
            return band

        def band_edge_center(values: Sequence[Sequence[Optional[float]]], start_row: int, end_row: int) -> Tuple[Optional[float], Optional[float]]:
            edge_values: List[float] = []
            center_values: List[float] = []
            cols = len(values[0]) if values else 0
            edge_cols = [0, cols - 1] if cols else []
            center_cols = list(range(max(0, cols // 2 - 2), min(cols, cols // 2 + 3)))
            for row in range(start_row - 1, end_row):
                if row < 0 or row >= len(values):
                    continue
                for col in edge_cols:
                    value = values[row][col]
                    if value is not None and math.isfinite(value):
                        edge_values.append(value)
                for col in center_cols:
                    value = values[row][col]
                    if value is not None and math.isfinite(value):
                        center_values.append(value)
            return (
                statistics.mean(edge_values) if edge_values else None,
                statistics.mean(center_values) if center_values else None,
            )

        def add_cells(row_index: int, cells: Sequence[str]) -> None:
            row_cells.setdefault(row_index, []).extend(cells)

        for sample_index, sample in enumerate(results):
            base_row = 1 + sample_index * (block_height + block_gap)
            title = sample_title(sample)
            grid_rows = min(REFERENCE_GRID_ROWS, sample.rows)
            grid_cols = min(REFERENCE_GRID_COLS, sample.cols)

            add_cells(
                base_row,
                [
                    self.cell(base_row, 2, title, 1),
                    self.cell(base_row, 4, "15x35 MPCD 결과", 1),
                ],
            )
            header_row = base_row + 2
            header_cells = [
                self.cell(header_row, 2, title, 1),
                self.cell(header_row, 3, "행", 2),
            ]
            for col in range(grid_cols):
                header_cells.append(self.cell(header_row, 4 + col, f"C{col + 1}", 2))
            header_cells.append(self.cell(header_row, 19, "구분", 1))
            header_cells.append(self.cell(header_row, 20, "값", 1))
            add_cells(header_row, header_cells)

            for row_offset in range(grid_rows):
                row_index = header_row + 1 + row_offset
                cells = []
                cells.append(self.cell(row_index, 3, row_offset + 1, 2))
                for col_offset in range(grid_cols):
                    value = sample.values[row_offset][col_offset]
                    cells.append(self.cell(row_index, 4 + col_offset, value, self.style_for_color(interpolate_color(value))))
                add_cells(row_index, cells)

            summary = sample.side_summary or compute_side_summary(sample.values)
            for idx, label in enumerate(summary_labels):
                row_index = base_row + 28 + idx
                value = summary[idx][1] if idx < len(summary) else None
                add_cells(
                    row_index,
                    [
                        self.cell(row_index, 19, label, 1),
                        self.cell(row_index, 20, value, 3),
                    ],
                )

            lower_header_row = base_row + 40
            lower_header = [self.cell(lower_header_row, 3, "행", 2)]
            for col in range(grid_cols):
                lower_header.append(self.cell(lower_header_row, 4 + col, f"C{col + 1}", 2))
            lower_header.append(self.cell(lower_header_row, 19, "edge", 1))
            lower_header.append(self.cell(lower_header_row, 20, "center", 1))
            add_cells(lower_header_row, lower_header)

            for idx, (label, band_label, start_row, end_row) in enumerate(
                [("6-10", "상", 6, 10), ("16-20", "중", 16, 20), ("26-30", "하", 26, 30)]
            ):
                row_index = lower_header_row + 1 + idx
                cells = [
                    self.cell(row_index, 2, label, 1),
                    self.cell(row_index, 3, band_label, 2),
                ]
                for col_offset, value in enumerate(band_means(sample.values, start_row, end_row)[:grid_cols]):
                    cells.append(self.cell(row_index, 4 + col_offset, value, self.style_for_color(interpolate_color(value))))
                edge, center = band_edge_center(sample.values, start_row, end_row)
                cells.append(self.cell(row_index, 19, edge, 3))
                cells.append(self.cell(row_index, 20, center, 3))
                add_cells(row_index, cells)

        max_row = max(1, len(results) * (block_height + block_gap) - block_gap)
        cols = [
            '<col min="2" max="2" width="8" customWidth="1"/>',
            '<col min="3" max="3" width="5.25" customWidth="1"/>',
            '<col min="4" max="18" width="3.5" customWidth="1"/>',
            '<col min="19" max="20" width="10" customWidth="1"/>',
        ]
        rows = [self.row(row_index, cells) for row_index, cells in sorted(row_cells.items())]
        drawing_xml = '<drawing r:id="rId1"/>' if results else ''
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="B1:T{max_row}"/>'
            '<sheetViews><sheetView workbookViewId="0" zoomScale="85" zoomScaleNormal="85"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16"/>'
            f'<cols>{"".join(cols)}</cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            f'{drawing_xml}'
            '</worksheet>'
        )

