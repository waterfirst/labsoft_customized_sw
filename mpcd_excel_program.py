#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
정리.xlsx의 MPCD 결과를 기준으로 동일한 색상/배치 Excel을 다시 생성하는 프로그램입니다.

기본 모드(reference_excel)는 폴더 안의 정리.xlsx에서 기존 MPCD 값을 읽어 옵니다.
따라서 원본 계산에 사용된 py 파일이 일부 누락되어 있어도 정리.xlsx와 같은 결과를
재현할 수 있습니다.

출력 배치 옵션:
  - excel_style: 정리.xlsx와 같은 6개 좌측 + 6개 우측 배치
  - vertical: 모든 시료를 아래 방향으로 연속 배치
  - by_time: 평가 시간 그룹별로 좌우/아래 배치
"""

from __future__ import annotations

import argparse
import csv
import errno
import json
import math
import os
import re
import shutil
import statistics
import struct
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import traceback
import zipfile
from dataclasses import dataclass
from datetime import UTC, datetime
from itertools import zip_longest
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET


COLOR_MIN_VALUE = -7.0
COLOR_MID_VALUE = 3.0
COLOR_MAX_VALUE = 13.0
COLOR_MIN = "F8696B"
COLOR_MID = "FFFFFF"
COLOR_MAX = "4D93D9"
RAW_GRID_ROWS = 15
RAW_GRID_COLS = 35
MPCD_UNIT = 0.0005
LUMINANCE_THRESHOLD = 300.0
DEFAULT_LABSOFT_SHORTCUT = Path(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TechnoTeam\Lmk LabSoft.lnk")
DEFAULT_LABSOFT_EXE = Path(r"C:\TechnoTeam\LabSoft\bin\lmk4.exe")
APP_CONFIG_DIR = Path(os.environ.get("APPDATA", Path.home())) / "MPCD_Excel_Program"
APP_CONFIG_PATH = APP_CONFIG_DIR / "settings.json"

REFERENCE_GRID_ROWS = 35
REFERENCE_GRID_COLS = 15
REFERENCE_BLOCK_STEP = 38
REFERENCE_LEFT_LABEL_COL = 1
REFERENCE_LEFT_GRID_COL = 2
REFERENCE_RIGHT_LABEL_COL = 20
REFERENCE_RIGHT_GRID_COL = 21
REFERENCE_FIRST_TOP_ROW = 2
SUMMARY_LABEL_OFFSET_COL = 16
SUMMARY_VALUE_OFFSET_COL = 17
SUMMARY_LABELS = [
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


def default_labsoft_path() -> str:
    if DEFAULT_LABSOFT_SHORTCUT.exists():
        return str(DEFAULT_LABSOFT_SHORTCUT)
    if DEFAULT_LABSOFT_EXE.exists():
        return str(DEFAULT_LABSOFT_EXE)
    return ""


def load_app_settings() -> Dict[str, str]:
    try:
        if APP_CONFIG_PATH.exists():
            data = json.loads(APP_CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {str(key): str(value) for key, value in data.items()}
    except Exception:
        pass
    return {}


def save_app_settings(settings: Dict[str, str]) -> None:
    APP_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    APP_CONFIG_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def read_json_file(path: Path) -> Dict[str, object]:
    try:
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


@dataclass
class SampleResult:
    sample_id: int
    values: List[List[Optional[float]]]
    side_summary: List[Tuple[str, Optional[float]]]
    source: str
    display_name: str = ""
    lxy_file: Optional[Path] = None
    wst_file: Optional[Path] = None

    @property
    def rows(self) -> int:
        return len(self.values)

    @property
    def cols(self) -> int:
        return len(self.values[0]) if self.values else 0


class OperationCancelled(RuntimeError):
    pass


def user_friendly_error(exc: Exception) -> Tuple[str, str]:
    raw = str(exc)
    if "LabSoft4 PCF export failed" in raw or "Open failed after retries" in raw:
        lmk_messages = {
            "LMK_PATH_EMPTY": "LMK LabSoft 경로가 비어 있습니다. 이 PC의 lmk4.exe 또는 Lmk LabSoft.lnk를 다시 선택하세요.",
            "LMK_SHORTCUT_FAILED": "LMK 바로가기(.lnk)를 읽지 못했습니다. 바로가기 대신 lmk4.exe를 직접 선택해 보세요.",
            "LMK_PATH_UNRESOLVED": "LMK 바로가기(.lnk)를 실제 실행 파일로 해석하지 못했습니다. lmk4.exe를 직접 선택해 보세요.",
            "LMK_PATH_NOT_FOUND": "선택된 LMK 경로가 현재 PC에 존재하지 않습니다. 다른 PC의 경로가 저장되어 있을 수 있습니다.",
            "LMK_START_FAILED": "LMK 실행 파일을 시작하지 못했습니다. 권한, 보안 프로그램 차단, 설치 상태를 확인하세요.",
            "LMK_COM_CREATE_FAILED": "LMK ActiveX/COM 객체를 만들지 못했습니다. LabSoft 설치 또는 COM 등록 문제일 가능성이 큽니다.",
            "LMK_OPEN_FAILED": "LMK 세션 열기(iOpen)에 실패했습니다. 라이선스, 최초 실행 팝업, 권한 차이, 수동 PCF 열기 가능 여부를 확인하세요.",
        }
        matched_message = ""
        matched_code = ""
        for code, message in lmk_messages.items():
            if f"code={code}" in raw:
                matched_code = code
                matched_message = message
                break
        cause_lines = [
            "LMK LabSoft 자동 실행/PCF 추출에 실패했습니다.",
            "",
            f"판단된 원인: {matched_message}" if matched_message else "판단된 원인: LMK 자동제어 연결 단계에서 실패했습니다.",
            f"오류 코드: {matched_code}" if matched_code else "",
            "",
            "가능성이 높은 원인:",
            "1. 이 PC에 LMK LabSoft 라이선스가 없거나 인증이 풀려 있습니다.",
            "2. 선택한 LMK 경로가 현재 PC의 실제 실행 파일과 다릅니다.",
            "3. LMK를 처음 실행할 때 뜨는 확인/라이선스 창 때문에 자동 연결이 막혔습니다.",
            "4. LMK ActiveX/COM 자동제어 등록이 깨졌거나 권한이 맞지 않습니다.",
            "5. 압축파일 안이나 임시 폴더에서 exe를 바로 실행했습니다.",
            "",
            "확인 순서:",
            "1. 압축을 일반 폴더에 완전히 푼 뒤 다시 실행하세요.",
            "2. LMK LabSoft를 수동 실행해 PCF 파일이 직접 열리는지 확인하세요.",
            "3. 프로그램의 LMK LabSoft 경로를 이 PC의 lmk4.exe 또는 바로가기(.lnk)로 다시 지정하세요.",
            "4. 계속 실패하면 LMK에서 Lxy/WST txt를 수동 추출한 뒤, 'PCF에서 새로 추출' 체크를 끄고 분석하세요.",
        ]
        if "rc=137363456" in raw or "0x08300000" in raw:
            cause_lines.insert(
                5,
                "   - 현재 오류코드 0x08300000은 보통 LMK 세션 열기/iOpen 실패로, 라이선스/초기 실행/COM 권한 문제일 가능성이 큽니다.",
            )
        return "LMK 자동 추출 실패", "\n".join(cause_lines)
    if isinstance(exc, FileNotFoundError):
        return "입력 파일 없음", raw
    return f"{type(exc).__name__}: {exc}", raw


@dataclass
class Placement:
    sample: SampleResult
    top_row: int
    label_col: int
    grid_col: int
    group_label: Optional[str] = None


@dataclass
class RawGeometry:
    width: int
    height: int
    sample_step: int
    cx: float
    cy: float
    cos: float
    sin: float
    u_min: float
    u_max: float
    v_min: float
    v_max: float
    tilt_deg: float


@dataclass
class RawCell:
    mpcd: Optional[float] = None


def discover_lxy_sample_ids(folder: Path) -> List[int]:
    ids = set()
    for path in folder.glob("*_Lxy.txt"):
        match = re.match(r"(\d+)_Lxy\.txt$", path.name, re.IGNORECASE)
        if match:
            ids.add(int(match.group(1)))
            continue
        stem = re.sub(r"_Lxy$", "", path.stem, flags=re.IGNORECASE)
        sample_id = pcf_sample_id(Path(stem))
        if sample_id is not None:
            ids.add(sample_id)
    return sorted(ids)


def pcf_sample_id(path: Path) -> Optional[int]:
    stem = path.stem.strip()
    match = re.match(r"^#?\s*(\d+)\s*(?:a+|w\b)?\s*$", stem, re.IGNORECASE)
    if match:
        return int(match.group(1))
    match = re.search(r"#\s*(\d+)(?:\D|$)", stem, re.IGNORECASE)
    return int(match.group(1)) if match else None


def pcf_group_key(path: Path) -> str:
    stem = path.stem.strip().lower()
    if "#" in stem:
        return stem.split("#", 1)[0].strip()
    return ""


def display_name_from_path(path: Path) -> str:
    stem = path.stem.strip()
    stem = re.sub(r"[_]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem)
    stem = re.sub(r"\s*#\s*", " #", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    return stem or path.stem or path.name


def sample_title(sample: SampleResult) -> str:
    return sample.display_name or f"Sample {sample.sample_id}"


def safe_output_stem(value: str, fallback: str) -> str:
    stem = (value or "").strip() or fallback
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem)
    stem = stem.rstrip(" .")
    return stem or fallback


def raw_txt_stems(sample_id: int, pcf_path: Optional[Path] = None) -> List[str]:
    stems: List[str] = []
    if pcf_path is not None:
        stems.append(safe_output_stem(pcf_path.stem, f"Sample {sample_id}"))
    stems.append(str(sample_id))
    unique: List[str] = []
    for stem in stems:
        if stem not in unique:
            unique.append(stem)
    return unique


def raw_txt_paths_for_stem(folder: Path, stem: str) -> Tuple[Path, Path]:
    return folder / f"{stem}_Lxy.txt", folder / f"{stem}_WST.txt"


def preferred_raw_txt_paths(folder: Path, sample_id: int, pcf_path: Optional[Path] = None) -> Tuple[Path, Path]:
    return raw_txt_paths_for_stem(folder, raw_txt_stems(sample_id, pcf_path)[0])


def existing_raw_txt_paths(folder: Path, sample_id: int, pcf_path: Optional[Path] = None) -> Tuple[Path, Path]:
    for stem in raw_txt_stems(sample_id, pcf_path):
        lxy_path, wst_path = raw_txt_paths_for_stem(folder, stem)
        if lxy_path.exists() and wst_path.exists():
            return lxy_path, wst_path
    for lxy_path in sorted(folder.glob("*_Lxy.txt"), key=natural_path_key):
        stem = re.sub(r"_Lxy$", "", lxy_path.stem, flags=re.IGNORECASE)
        if pcf_sample_id(Path(stem)) != sample_id:
            continue
        wst_path = folder / f"{stem}_WST.txt"
        if wst_path.exists():
            return lxy_path, wst_path
    return preferred_raw_txt_paths(folder, sample_id, pcf_path)


def metadata_path_for_raw(lxy_path: Path) -> Path:
    stem = lxy_path.name
    if stem.lower().endswith("_lxy.txt"):
        stem = stem[:-8]
    else:
        stem = lxy_path.stem
    return lxy_path.with_name(f"{stem}_MPCD_source.json")


def rotate_matrix_180(matrix: Sequence[Sequence[Optional[float]]]) -> List[List[Optional[float]]]:
    return [list(reversed(row)) for row in reversed(matrix)]


def apply_hole_mask(matrix: Sequence[Sequence[Optional[float]]]) -> List[List[Optional[float]]]:
    masked = [list(row) for row in matrix]
    if masked and len(masked[0]) >= 8:
        masked[0][7] = None
    return masked


def pcf_looks_cropped(path: Path) -> bool:
    stem = path.stem.strip().casefold()
    return bool(
        re.match(r"^#?\s*\d+\s*a+$", stem)
        or re.match(r"^#?\s*\d+\D+", stem) and not re.match(r"^#?\s*\d+\s*w\b", stem)
    )


def pcf_priority(path: Path) -> Tuple[int, int, str]:
    stem = path.stem.strip().casefold()
    name = path.name.casefold()
    if re.match(r"^\d+\s*a+$", stem):
        return (0, len(name), name)
    if re.match(r"^#\s*\d+\s*a+$", stem):
        return (1, len(name), name)
    if re.match(r"^\d+\D+", stem) and not re.match(r"^\d+\s*w\b", stem):
        return (2, len(name), name)
    if re.match(r"^\d+\s*w\b", stem):
        return (3, len(name), name)
    if re.match(r"^#\s*\d+\s*w\b", stem):
        return (4, len(name), name)
    if re.match(r"^#\s*\d+$", stem):
        return (5, len(name), name)
    return (6, len(name), name)


def natural_path_key(path: Path) -> Tuple[object, ...]:
    parts = re.split(r"(\d+)", path.name.casefold())
    return tuple(int(part) if part.isdigit() else part for part in parts)


def assign_manual_pcf_files(paths: Sequence[Path]) -> Dict[int, Path]:
    selected: Dict[int, Path] = {}
    next_id = 1
    for path in paths:
        sample_id = pcf_sample_id(path)
        if sample_id is None or sample_id in selected:
            while next_id in selected:
                next_id += 1
            sample_id = next_id
        selected[sample_id] = path
        next_id = max(next_id, sample_id + 1)
    return dict(sorted(selected.items()))


def discover_pcf_files(folder: Path, manual_pcf_files: Optional[Sequence[Path]] = None) -> Dict[int, Path]:
    if manual_pcf_files:
        return assign_manual_pcf_files([Path(path) for path in manual_pcf_files if Path(path).is_file()])
    paths: List[Path] = []
    for pattern in ("*.pcf", "*.pct"):
        for path in folder.glob(pattern):
            if path.is_file():
                paths.append(path)
    grouped: Dict[int, List[Path]] = {}
    unnumbered: List[Path] = []
    for path in paths:
        sample_id = pcf_sample_id(path)
        if sample_id is None:
            unnumbered.append(path)
        else:
            grouped.setdefault(sample_id, []).append(path)

    selected: Dict[int, Path] = {
        sample_id: sorted(candidates, key=pcf_priority)[0]
        for sample_id, candidates in grouped.items()
    }
    next_id = max(selected, default=0) + 1
    for path in sorted(unnumbered, key=natural_path_key):
        while next_id in selected:
            next_id += 1
        selected[next_id] = path
        next_id += 1
    return dict(sorted(selected.items()))


def discover_sample_ids(folder: Path, manual_pcf_files: Optional[Sequence[Path]] = None) -> List[int]:
    ids = set(discover_lxy_sample_ids(folder))
    ids.update(discover_pcf_files(folder, manual_pcf_files))
    return sorted(ids)


def suggest_reference_excel(folder: Path) -> Path:
    folder = Path(folder)
    folder_name = folder.name.strip()
    time_label = folder_name if folder_name.lower().endswith("hr") else f"{folder_name}hr"
    time_specific_names = [
        f"정리 - {time_label}.xlsx",
        f"정리-{time_label}.xlsx",
        f"정리_{time_label}.xlsx",
        f"정리 {time_label}.xlsx",
    ]

    candidates: List[Path] = []
    candidates.append(folder / "정리.xlsx")
    for name in time_specific_names:
        candidates.append(folder / name)
    for name in time_specific_names:
        candidates.append(folder.parent / name)
    for base in (folder, folder.parent):
        if not base.exists():
            continue
        candidates.extend(
            sorted(
                path
                for path in base.glob("*.xlsx")
                if not path.name.startswith("~$")
                and not path.name.startswith("MPCD_")
                and "정리" in path.name
                and (folder_name in path.name or time_label in path.name)
            )
        )
    candidates.append(folder.parent / "정리.xlsx")

    seen = set()
    for candidate in candidates:
        key = str(candidate.resolve()) if candidate.exists() else str(candidate)
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists():
            return candidate
    return folder / "정리.xlsx"


def resolve_reference_excel(folder: Path, reference_path: Optional[Path] = None) -> Path:
    if reference_path and reference_path.exists():
        return reference_path

    suggested = suggest_reference_excel(folder)
    if suggested.exists():
        return suggested

    attempted = reference_path or suggested
    raise FileNotFoundError(
        "기준 Excel을 찾지 못했습니다.\n"
        f"- Data 폴더: {folder}\n"
        f"- 찾으려던 파일: {attempted}\n"
        "Data 폴더 안에 정리.xlsx를 넣거나, 상위 폴더에 '정리 - 360hr.xlsx'처럼 시간명이 포함된 정리 파일을 두세요.\n"
        "또는 프로그램의 '기준 Excel'에서 파일을 직접 선택해 주세요."
    )


def automatic_output_path(folder: Path) -> Path:
    folder = Path(folder)
    label = folder.name.strip() or "data"
    safe_label = re.sub(r'[<>:"/\\|?*]+', "_", label)
    return folder / f"MPCD_{safe_label}_results.xlsx"


def available_output_path(output_path: Path) -> Path:
    if not output_path.exists():
        return output_path
    for index in range(1, 1000):
        candidate = output_path.with_name(f"{output_path.stem}_{index}{output_path.suffix}")
        if not candidate.exists():
            return candidate
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_path.with_name(f"{output_path.stem}_{timestamp}{output_path.suffix}")


LABSOFT_EXPORT_SCRIPT = r"""
param(
    [Parameter(Mandatory=$true)][string]$PairsPath,
    [Parameter(Mandatory=$true)][string]$DumpRoot,
    [string]$LabSoftPath = ''
)
$ErrorActionPreference = 'Stop'
function CheckRc($Name, $Rc) {
    if ($Rc -ne 0) {
        $hex = ('0x{0:X8}' -f ([uint32]$Rc))
        throw "$Name failed. rc=$Rc ($hex)"
    }
}
function FailUser($Code, $Message) {
    throw "MPCD_LMK_ERROR code=$Code message=$Message"
}
function NewGreyImage($Lab, $Name) {
    $index = 0
    CheckRc "ImageCreate $Name" ($Lab.iImageCreate(0, $Name, [ref]$index))
    return $index
}
function ResolveShortcutInfo($PathText) {
    $info = [ordered]@{
        Path = ''
        Arguments = ''
        WorkingDirectory = ''
        Source = $PathText
    }
    if ([string]::IsNullOrWhiteSpace($PathText)) { return [pscustomobject]$info }
    if (!(Test-Path -LiteralPath $PathText)) {
        $info.Path = $PathText
        return [pscustomobject]$info
    }
    if ([System.IO.Path]::GetExtension($PathText).ToLowerInvariant() -eq '.lnk') {
        try {
            $ws = New-Object -ComObject WScript.Shell
            $shortcut = $ws.CreateShortcut($PathText)
            $info.Path = $shortcut.TargetPath
            $info.Arguments = $shortcut.Arguments
            $info.WorkingDirectory = $shortcut.WorkingDirectory
        } catch {
            FailUser "LMK_SHORTCUT_FAILED" "LMK shortcut could not be read. Path='$PathText'. Error='$($_.Exception.Message)'."
        }
        return [pscustomobject]$info
    }
    $info.Path = $PathText
    $info.WorkingDirectory = Split-Path -Parent $PathText
    return [pscustomobject]$info
}
$labsoftInfo = ResolveShortcutInfo $LabSoftPath
$resolvedLabSoft = [string]$labsoftInfo.Path
if ([string]::IsNullOrWhiteSpace($LabSoftPath)) {
    FailUser "LMK_PATH_EMPTY" "LMK LabSoft path is empty. Select lmk4.exe or Lmk LabSoft.lnk."
}
if ([string]::IsNullOrWhiteSpace($resolvedLabSoft)) {
    FailUser "LMK_PATH_UNRESOLVED" "LMK shortcut could not be resolved. LabSoftPath='$LabSoftPath'."
}
if (!(Test-Path -LiteralPath $resolvedLabSoft)) {
    FailUser "LMK_PATH_NOT_FOUND" "Resolved LMK executable was not found. LabSoftPath='$LabSoftPath' ResolvedPath='$resolvedLabSoft'."
}
if (![string]::IsNullOrWhiteSpace($resolvedLabSoft) -and (Test-Path -LiteralPath $resolvedLabSoft)) {
    try {
        $processName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedLabSoft)
        $isRunning = $false
        foreach ($process in Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            try {
                if ($process.Path -eq $resolvedLabSoft) { $isRunning = $true; break }
            } catch {
                $isRunning = $true
                break
            }
        }
        if (!$isRunning) {
            $startArgs = @{
                FilePath = $resolvedLabSoft
                WindowStyle = 'Minimized'
            }
            if (![string]::IsNullOrWhiteSpace($labsoftInfo.Arguments)) {
                $startArgs.ArgumentList = $labsoftInfo.Arguments
            }
            if (![string]::IsNullOrWhiteSpace($labsoftInfo.WorkingDirectory) -and (Test-Path -LiteralPath $labsoftInfo.WorkingDirectory)) {
                $startArgs.WorkingDirectory = $labsoftInfo.WorkingDirectory
            }
            Start-Process @startArgs | Out-Null
            Start-Sleep -Milliseconds 8000
        }
    } catch {
        FailUser "LMK_START_FAILED" "LMK executable could not be started. Path='$resolvedLabSoft'. Error='$($_.Exception.Message)'."
    }
}
function OpenLabSoftSession() {
    $lastError = ''
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            try {
                $candidate = New-Object -ComObject 'lmk4.LmkAxServer.1'
            } catch {
                FailUser "LMK_COM_CREATE_FAILED" "LMK ActiveX/COM object could not be created. Check LabSoft installation and COM registration. Error='$($_.Exception.Message)'."
            }
            $rc = $candidate.iOpen()
            if ($rc -eq 0) {
                return $candidate
            }
            $hex = ('0x{0:X8}' -f ([uint32]$rc))
            $lastError = "attempt=$attempt rc=$rc ($hex)"
        } catch {
            $lastError = "attempt=$attempt $($_.Exception.Message)"
        }
        Start-Sleep -Milliseconds 2000
    }
    FailUser "LMK_OPEN_FAILED" "LMK session open failed after retries. $lastError. LabSoftPath='$LabSoftPath' ResolvedPath='$resolvedLabSoft'. Check license, first-run popup, manual PCF open, and permission level."
}
$lab = OpenLabSoftSession
try { $lab.iShow(2) | Out-Null } catch {}
$pairs = Import-Csv -Delimiter "`t" -Path $PairsPath
foreach ($pair in $pairs) {
    $sample = [string]$pair.sample_id
    $pcf = [string]$pair.pcf_path
    if (!(Test-Path -LiteralPath $pcf)) {
        throw "Sample $sample PCF file not found: $pcf"
    }
    Write-Output "BEGIN`t$sample`t$pcf"
    CheckRc "LoadImage $sample" ($lab.iLoadImage(-1, $pcf))
    $firstLine = 0
    $lastLine = 0
    $firstColumn = 0
    $lastColumn = 0
    $dimensions = 0
    CheckRc "ImageGetSize $sample" ($lab.iImageGetSize(-1, [ref]$firstLine, [ref]$lastLine, [ref]$firstColumn, [ref]$lastColumn, [ref]$dimensions))
    if ($dimensions -ne 3) {
        throw "Sample $sample is not a 3-channel color image. dimensions=$dimensions"
    }
    $sampleDir = Join-Path $DumpRoot $sample
    New-Item -ItemType Directory -Force -Path $sampleDir | Out-Null
    $l1 = NewGreyImage $lab "mpcd_lxy_${sample}_1"
    $l2 = NewGreyImage $lab "mpcd_lxy_${sample}_2"
    $l3 = NewGreyImage $lab "mpcd_lxy_${sample}_3"
    CheckRc "SplitColorImage Lxy $sample" ($lab.iSplitColorImage(-1, $l1, $l2, $l3, 32))
    CheckRc "Dump Lxy1 $sample" ($lab.iImageGetDumpToFile($l1, $firstLine, $lastLine, $firstColumn, $lastColumn, (Join-Path $sampleDir 'lxy_1.bin')))
    CheckRc "Dump Lxy2 $sample" ($lab.iImageGetDumpToFile($l2, $firstLine, $lastLine, $firstColumn, $lastColumn, (Join-Path $sampleDir 'lxy_2.bin')))
    CheckRc "Dump Lxy3 $sample" ($lab.iImageGetDumpToFile($l3, $firstLine, $lastLine, $firstColumn, $lastColumn, (Join-Path $sampleDir 'lxy_3.bin')))
    $w1 = NewGreyImage $lab "mpcd_wst_${sample}_1"
    $w2 = NewGreyImage $lab "mpcd_wst_${sample}_2"
    $w3 = NewGreyImage $lab "mpcd_wst_${sample}_3"
    CheckRc "SplitColorImage WST $sample" ($lab.iSplitColorImage(-1, $w1, $w2, $w3, 16384))
    CheckRc "Dump WST1 $sample" ($lab.iImageGetDumpToFile($w1, $firstLine, $lastLine, $firstColumn, $lastColumn, (Join-Path $sampleDir 'wst_1.bin')))
    CheckRc "Dump WST2 $sample" ($lab.iImageGetDumpToFile($w2, $firstLine, $lastLine, $firstColumn, $lastColumn, (Join-Path $sampleDir 'wst_2.bin')))
    CheckRc "Dump WST3 $sample" ($lab.iImageGetDumpToFile($w3, $firstLine, $lastLine, $firstColumn, $lastColumn, (Join-Path $sampleDir 'wst_3.bin')))
    foreach ($index in @($l1, $l2, $l3, $w1, $w2, $w3)) {
        try { $lab.iImageDelete($index) | Out-Null } catch {}
    }
    Write-Output "DONE`t$sample`t$firstLine`t$lastLine`t$firstColumn`t$lastColumn`t$dimensions"
}
$lab = $null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
"""


def load_float_dump(path: Path, count: int) -> Tuple[float, ...]:
    data = path.read_bytes()
    expected_size = count * 4
    if len(data) != expected_size:
        raise ValueError(f"{path.name}: dump size mismatch. expected {expected_size}, got {len(data)}")
    return struct.unpack(f"<{count}f", data)


def write_labsoft_txt(channels: Sequence[Tuple[float, ...]], rows: int, cols: int, output_path: Path) -> None:
    if len(channels) != 3:
        raise ValueError("LabSoft export requires exactly 3 channels.")
    with output_path.open("w", encoding="ascii", newline="\n") as handle:
        for row in range(rows):
            base = row * cols
            values: List[str] = []
            for col in range(cols):
                index = base + col
                values.extend(f"{channel[index]:.3e}" for channel in channels)
            handle.write("\t".join(values))
            handle.write("\n")


def detect_luminance_crop_bounds(
    luminance_path: Path,
    rows: int,
    cols: int,
    threshold: float = LUMINANCE_THRESHOLD,
) -> Optional[Tuple[int, int, int, int]]:
    row_bytes = cols * 4
    row_format = f"<{cols}f"
    first_row = rows
    last_row = -1
    first_col = cols
    last_col = -1
    with luminance_path.open("rb") as handle:
        for row_index in range(rows):
            row_data = struct.unpack(row_format, handle.read(row_bytes))
            hit_cols = [col for col, value in enumerate(row_data) if math.isfinite(value) and value >= threshold]
            if not hit_cols:
                continue
            first_row = min(first_row, row_index)
            last_row = max(last_row, row_index)
            first_col = min(first_col, min(hit_cols))
            last_col = max(last_col, max(hit_cols))
    if last_row < first_row or last_col < first_col:
        return None
    return first_row, last_row, first_col, last_col


def write_labsoft_txt_from_dumps(
    channel_paths: Sequence[Path],
    rows: int,
    cols: int,
    output_path: Path,
    bounds: Optional[Tuple[int, int, int, int]] = None,
) -> None:
    if len(channel_paths) != 3:
        raise ValueError("LabSoft export requires exactly 3 dump files.")
    expected_size = rows * cols * 4
    for path in channel_paths:
        actual_size = path.stat().st_size
        if actual_size != expected_size:
            raise ValueError(f"{path.name}: dump size mismatch. expected {expected_size}, got {actual_size}")

    if bounds is None:
        row_start, row_end, col_start, col_end = 0, rows - 1, 0, cols - 1
    else:
        row_start, row_end, col_start, col_end = bounds
        row_start = max(0, min(rows - 1, row_start))
        row_end = max(row_start, min(rows - 1, row_end))
        col_start = max(0, min(cols - 1, col_start))
        col_end = max(col_start, min(cols - 1, col_end))

    row_bytes = cols * 4
    row_format = f"<{cols}f"
    with output_path.open("w", encoding="ascii", newline="\n") as output_handle:
        with channel_paths[0].open("rb") as first, channel_paths[1].open("rb") as second, channel_paths[2].open("rb") as third:
            for handle in (first, second, third):
                handle.seek(row_start * row_bytes)
            for _ in range(row_start, row_end + 1):
                rows_data = [
                    struct.unpack(row_format, handle.read(row_bytes))
                    for handle in (first, second, third)
                ]
                values: List[str] = []
                for col in range(col_start, col_end + 1):
                    values.extend(f"{channel[col]:.3e}" for channel in rows_data)
                output_handle.write("\t".join(values))
                output_handle.write("\n")


def parse_labsoft_done_lines(output: str) -> Dict[int, Tuple[int, int]]:
    sizes: Dict[int, Tuple[int, int]] = {}
    for line in output.splitlines():
        parts = line.strip().split("\t")
        if len(parts) == 7 and parts[0] == "DONE":
            sample_id = int(parts[1])
            first_line = int(parts[2])
            last_line = int(parts[3])
            first_col = int(parts[4])
            last_col = int(parts[5])
            sizes[sample_id] = (last_line - first_line + 1, last_col - first_col + 1)
    return sizes


def run_labsoft_dump(
    pairs: Sequence[Tuple[int, Path]],
    dump_root: Path,
    progress: Optional[callable] = None,
    labsoft_path: str = "",
) -> Dict[int, Tuple[int, int]]:
    with tempfile.TemporaryDirectory(prefix="mpcd_labsoft_") as temp_name:
        temp_dir = Path(temp_name)
        pairs_path = temp_dir / "pairs.tsv"
        script_path = temp_dir / "export_labsoft.ps1"
        with pairs_path.open("w", encoding="utf-8", newline="\n") as handle:
            handle.write("sample_id\tpcf_path\n")
            for sample_id, pcf_path in pairs:
                handle.write(f"{sample_id}\t{pcf_path.resolve()}\n")
        script_path.write_text(LABSOFT_EXPORT_SCRIPT, encoding="utf-8")
        def make_command(powershell_path: str) -> List[str]:
            command = [
                powershell_path,
                "-NoProfile",
                "-Sta",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-PairsPath",
                str(pairs_path),
                "-DumpRoot",
                str(dump_root),
            ]
            if labsoft_path:
                command.extend(["-LabSoftPath", labsoft_path])
            return command

        powershell_candidates = ["powershell"]
        if os.name == "nt":
            syswow64_powershell = (
                Path(os.environ.get("SystemRoot", r"C:\Windows"))
                / "SysWOW64"
                / "WindowsPowerShell"
                / "v1.0"
                / "powershell.exe"
            )
            if syswow64_powershell.exists():
                powershell_candidates.append(str(syswow64_powershell))
        if progress:
            progress(f"LabSoft4 PCF export: {len(pairs)} sample(s)...")
        failures: List[str] = []
        for attempt, powershell_path in enumerate(powershell_candidates, start=1):
            if progress and attempt > 1:
                progress(f"LabSoft4 PCF export retry with {Path(powershell_path).name}...")
            completed = subprocess.run(
                make_command(powershell_path),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if completed.returncode == 0:
                return parse_labsoft_done_lines(completed.stdout)
            details = "\n".join(part for part in (completed.stdout, completed.stderr) if part.strip())
            failures.append(f"[{powershell_path}]\n{details}")
        raise RuntimeError("LabSoft4 PCF export failed.\n" + "\n\n".join(failures))


