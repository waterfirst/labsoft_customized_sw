# MPCD Excel Program 사용법

## 포함 파일

- `dist\MPCD_Excel_Program.exe`
  - 배포용 GUI 실행 파일입니다.
  - Python 설치 없이 실행할 수 있습니다.
- `mpcd_excel_program.py`
  - 원본 Python 프로그램입니다.
- `README_MPCD.md`
  - 사용법과 주의사항 문서입니다.

## 입력 파일

프로그램은 아래 두 가지 입력 방식을 지원합니다.

1. 기존 LabSoft export txt 사용
   - `1_Lxy.txt`, `1_WST.txt`
   - `2_Lxy.txt`, `2_WST.txt`
   - 같은 규칙으로 시료 번호별 파일을 둡니다.

2. LabSoft4 PCF 자동 추출 사용
   - 파일명 제약 없이 폴더 안의 모든 `.pcf`, `.pct` 파일을 인식합니다.
   - 파일명 자연 정렬 순서대로 `Sample 1`, `Sample 2`, `Sample 3`...을 자동 부여합니다.
   - 기본 설정에서는 PCF가 있으면 기존 `N_Lxy.txt`, `N_WST.txt`가 있어도 LabSoft4 ActiveX로 새로 추출한 뒤 MPCD Excel을 만듭니다.
   - 생성된 `N_Lxy.txt`, `N_WST.txt`는 Data 폴더에 저장됩니다.

## LabSoft4 자동 추출 조건

- 이 PC에는 LabSoft4가 아래 경로에 설치되어 있어야 합니다.
  - `C:\TechnoTeam\LabSoft\bin\lmk4.exe`
- LabSoft4 ActiveX 객체가 등록되어 있어야 합니다.
  - 확인된 객체명: `lmk4.LmkAxServer.1`
- 자동 추출은 Windows에서만 동작합니다.
- 자동 추출 중 LabSoft4가 잠깐 실행되거나 최소화 상태로 뜰 수 있습니다.
- 배포받은 PC에서는 GUI의 `LMK LabSoft` 행에서 본인 PC의 `Lmk LabSoft.lnk` 또는 `lmk4.exe`를 선택할 수 있습니다.
- 선택한 LMK 경로는 사용자 설정 폴더에 저장되어 다음 실행 때 자동으로 불러옵니다.

## 기본 실행 방법

1. `dist\MPCD_Excel_Program.exe`를 실행합니다.
2. `Panel ROI`를 필요에 따라 입력합니다.
   - 기본 ROI는 `0.98`입니다.
3. PCF 자동 추출을 사용할 경우 `LMK LabSoft`에서 본인 PC의 LabSoft 바로가기 또는 실행 파일을 선택합니다.
4. `Data 입력`에서 분석할 폴더를 선택합니다.
   - 기본은 폴더 안 PCF에서 Lxy/WST를 새로 추출해 계산합니다.
   - `PCF에서 새로 추출` 체크를 끄면 기존 txt를 우선 사용합니다.
5. 분석 시료 선택에서 시료와 배치 방식을 선택합니다. 시료 목록은 폴더 안의 PCF/TXT 개수에 맞춰 표시됩니다.
   - 시료를 선택하지 않으면 전체 시료를 분석합니다.
   - 여러 시료는 `Ctrl` 또는 `Shift`로 복수 선택할 수 있습니다.
   - `전체 선택` 버튼을 누르면 현재 목록의 모든 시료를 선택합니다.
6. `분석` 버튼을 누릅니다.
7. 결과 Excel과 CSV는 선택한 Data 폴더에 자동 생성됩니다.

예를 들어 `360` 폴더를 선택하면 다음과 같은 결과가 생성됩니다.

- `MPCD_360_results.xlsx`
- `MPCD_360_results.csv`

같은 이름의 결과 파일이 이미 있으면 `_1`, `_2`를 붙여 새 파일로 저장합니다.

## 결과 파일

- 첫 번째 sheet: `새로운 결과1`
  - 시료 1개당 하나의 결과 블록으로 출력합니다.
  - 정리 Excel의 `새로운 결과1` 형식을 기준으로 맞췄습니다.
  - 시료명은 `Sample 1` 대신 입력 PCF/TXT 파일명에서 가져온 이름을 표시합니다.
  - 각 시료 블록 아래에 상/중/하 평균 프로파일 그래프를 포함합니다.
- `MPCD Map`
  - 전체 시료 MPCD map을 배치 옵션에 따라 보여줍니다.
- `Graph`
  - 시료별 Row Profile, Column Profile, Summary Graph 데이터를 색상 스케일과 함께 보여줍니다.
- `All Data`
  - 전체 MPCD 수치 데이터입니다.
- `Summary`
  - 입력 파일, layout, 생성 시각 등 요약 정보입니다.
- CSV
  - `Sample, Row, Col, MPCD` 형식입니다.

## 분석 시료 선택

- `정리 Excel식`
  - 기존 정리 파일과 비슷하게 1~6번은 왼쪽, 7~12번은 오른쪽으로 배치합니다.
  - 12개를 초과해도 다음 묶음으로 이어서 출력합니다.
- `전체 아래로`
  - 모든 시료를 한 줄 흐름으로 아래 방향에 이어서 출력합니다.
- 결과 data는 hole부가 상단에 오도록 배열됩니다.

## 색상 기준

MPCD 값은 아래 기준으로 색 표현합니다.

- 최소값 `-7`: `#F8696B` 빨강
- 중간값 `3`: `#FFFFFF` 흰색
- 최대값 `13`: `#4D93D9` 파랑

## 주의사항

- PCF 자동 추출은 LabSoft4가 설치된 PC에서만 가능합니다.
- PCF 파일명에는 제약이 없습니다. 다만 결과의 Sample 번호는 파일명 자연 정렬 순서로 자동 부여됩니다.
- 기본 설정에서는 PCF가 있으면 `N_Lxy.txt`, `N_WST.txt`를 새로 생성해 덮어씁니다. 기존 txt를 보존해서 계산하려면 `PCF에서 새로 추출` 체크를 끄세요.
- 이미 `N_Lxy.txt`, `N_WST.txt`가 있으면 실행 시 재생성 여부를 묻습니다. `아니오`를 선택하면 가능한 기존 txt로 분석합니다.
- PCF에서 생성된 `N_Lxy.txt`, `N_WST.txt`는 Data 폴더에 남겨 둡니다.
- PCF만 있는 폴더에서는 자동 추출 시간이 추가로 걸립니다.
- 결과 Excel을 열어둔 상태에서 다시 실행하면 같은 이름으로 저장하지 않고 번호가 붙은 새 파일을 만듭니다.
- 오류가 발생하면 Data 입력 폴더에 `MPCD_error_log.txt`가 생성됩니다.

## 명령줄 실행 예시

```bat
MPCD_Excel_Program.exe --no-gui --folder "D:\data\360" --layout excel_style --roi 0.98 --source pcf
python mpcd_excel_program.py --no-gui --folder ".\360" --layout vertical
python mpcd_excel_program.py --no-gui --folder ".\360" --layout by_time --time-groups "0hr,240hr" --source txt
```

## 배포 권장 구성

다른 PC에 전달할 때는 아래 파일을 함께 전달하는 것을 권장합니다.

```text
dist\MPCD_Excel_Program.exe
README_MPCD.md
데이터 폴더
```

데이터 폴더에는 `*_Lxy.txt`/`*_WST.txt` 또는 `*W.pcf` 파일이 들어 있으면 됩니다.
