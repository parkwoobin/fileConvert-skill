# FileConvert Skill

HWPX, HWP, Word, PDF, PowerPoint 문서를 로컬에서 양방향으로 변환하고 렌더링 결과를 비교하는 Codex 스킬입니다.

| 양방향 | 변환 엔진 |
| --- | --- |
| HWPX ↔ DOCX | HWPX XML의 본문과 표를 Word 문서로 구성합니다. DOCX에서 시작하면 한글 문서로 저장합니다. |
| HWP ↔ DOCX | HWP 5.0의 본문과 표를 Word 문서로 구성합니다. DOCX에서 시작하면 한글 문서로 저장합니다. |
| PDF ↔ HWPX, PDF ↔ HWP | PDF의 내용과 배치를 한글 문서로 가져옵니다. 한글 문서에서 시작하면 PDF로 내보냅니다. |
| PDF ↔ DOCX, PDF ↔ DOC | PDF의 본문과 배치를 Word 문서로 재구성합니다. Word 문서에서 시작하면 PDF로 내보냅니다. |
| PDF ↔ PPTX, PDF ↔ PPT | PowerPoint PDF 내보내기. PDF에서 시작하면 페이지별 이미지를 슬라이드에 배치 |

Windows 데스크톱, Python 3.13(검증 버전)과 경로별로 설치·활성화된 한컴오피스 한글 및 Microsoft Word·PowerPoint가 필요합니다. PDF → PPTX 후보 생성과 HWPX 구조 추출은 Python으로 처리하지만 전체 CLI 검증에는 해당 오피스 앱이 필요합니다. Hancom 파일 접근 대화상자가 뜨는 환경에서는 한컴 공식 문서에 따라 등록한 경로 확인 모듈 이름을 `FILECONVERT_HANCOM_PATH_MODULE`에 지정할 수 있습니다.

```powershell
python -m pip install -r requirements.txt
python scripts/convert.py --list
python scripts/convert.py input.hwpx output.docx --report comparison.json
```

기본 모드는 각 페이지의 크기와 144 DPI RGB 픽셀을 비교합니다. 페이지 수나 크기 또는 픽셀이 하나라도 다르면 종료 코드 2를 반환하고 결과 문서 저장을 보류합니다. 비교 보고서에는 페이지별 크기와, 픽셀 크기가 같을 경우 달라진 영역의 경계가 기록됩니다. 페이지 수가 같을 때는 `--allow-differences`로 외형이 달라진 후보를 저장할 수 있습니다. 페이지 수가 달라진 편집용 문서를 받으려면 `--allow-page-count-change`도 추가하세요. PDF에서 HWP/HWPX로 변환할 때는 이 옵션을 사용해도 원본과 같은 페이지 수를 요구합니다. 코드 1은 변환 또는 검증 실패입니다. 동일 판정은 실행 환경과 렌더링 설정에 한정됩니다. 글꼴 파일과 편집 가능한 수식·각주·미주·필드의 보존은 대상 앱에서 별도로 확인해야 합니다.

HWPX에서 편집 가능한 본문을 먼저 얻고 싶으면 `--hwpx-structure`를 사용합니다. OWPML의 섹션, 문자 속성, 그림 참조, 표 및 셀 병합을 직접 읽습니다. 대표적인 글꼴·크기·색상과 용지 크기·여백을 매핑합니다. 모든 언어의 자간 값이 동일하면 Word 단위로 근사합니다. 언어별 자간·정확한 배치·수식·각주·미주·머리말·꼬리말은 이 경로에서 보존되지 않습니다. 이 옵션도 기본 비교를 거칩니다.

PDF를 Word나 한글 문서로 가져올 때 단, 머리말, 각주, 수식 및 줄바꿈이 재구성되는 방식은 원본과 다를 수 있습니다. 이미지로 된 PDF에 별도 OCR은 적용하지 않습니다. PowerPoint 경로는 페이지 이미지를 사용하므로 슬라이드의 텍스트는 개별 편집할 수 없습니다. 서로 다른 크기의 PDF 페이지는 첫 페이지 크기에 맞춰 여백을 추가할 수 있습니다.

Codex에 설치하려면 이 저장소를 사용자 스킬 폴더 아래 `file-convert-skill` 폴더로 복제하세요. 한컴과 Microsoft Office 자동화는 데스크톱 사용을 전제로 합니다.

참고한 방식은 HWP 바이너리와 HWPX XML 경로를 구분한 [morningkim/hwp2docx](https://github.com/morningkim/hwp2docx)와 HWPX의 문단·표를 읽어 DOCX를 구성한 [kangdacool/hwpx-editing-skill](https://github.com/kangdacool/hwpx-editing-skill/blob/main/skills/hwpx-editing/scripts/hwpx_to_docx.py)입니다. 기본 HWP/HWPX 변환기는 아래 고정 버전의 소스를 포함하며, 함께 제공되는 MIT 라이선스를 따릅니다. PDF → DOCX에는 ArtifexSoftware의 pdf2docx 0.5.13을 사용합니다. 한컴 형식 처리는 [한컴 자동화 문서](https://developer.hancom.com/en-us/hwpautomation)를 참고했습니다.

## 검증 결과

Windows에서 합성 문서로 16개 방향 모두 결과 파일 생성을 확인했습니다. PDF로 내보내는 6개 방향은 해당 샘플의 렌더 비교를 통과했습니다. 나머지 10개 방향의 샘플은 외형 차이가 있어 엄격 모드 통과 대상으로 볼 수 없습니다.

별도의 Word → HWPX → Word 왕복 샘플에서 수식·각주·미주·머리말·꼬리말·페이지 번호의 문서 구조는 남았습니다. 자간 1.25pt는 1.20pt로 바뀌었고, 첫 한글 가져오기에서 인라인 그림이 누락됐습니다. 이 검증 결과는 문서마다 달라질 수 있습니다.

## PDF의 글자와 도형을 Word 개체로 분리

고정 배치 PDF에는 `python scripts/convert.py input.pdf output.docx --pdf-objects --allow-differences --report report.json`을 사용할 수 있습니다. 화면 위치에 맞춰 수정 가능한 텍스트 상자, Word 도형, 개별 그림을 배치합니다. 같은 흐름의 여러 줄은 한 텍스트 상자로 묶고, 원본 렌더링을 기준으로 숨은 텍스트와 근접 중복을 걸러냅니다. 자간과 폰트는 근사치이며 해당 글꼴이 설치되어 있어야 원본에 가깝게 표시됩니다. 일부 PDF 그림은 잘라낸 이미지로 들어갑니다. 클리핑, 겹침, 그룹의 복합 효과는 부분적으로만 처리하고 PDF의 표와 문단 구조를 재현하지 않습니다. 개체 편집 검증은 Windows용 Microsoft Word에서 했습니다.

## 겟컨버트 웹 데모

이 저장소에서 의존성을 설치한 뒤 다음 명령을 실행하고 브라우저에서 `http://127.0.0.1:8765/`를 엽니다.

```powershell
python -m pip install -r requirements.txt
python demo/server.py
```

파일 확장자에 맞는 출력 형식을 선택합니다. 편집 가능한 문서로 변환하는 경로를 출력 형식에 맞게 자동 선택합니다. PDF → DOCX에는 pdf2docx를 사용해 문단과 실제 Word 표를 재구성합니다. 이미지로만 된 PDF를 Word나 한글 문서로 변환하려면 먼저 OCR이 필요합니다. 편집 가능한 문서 중 일부는 변환 후 페이지 수가 달라져도 원본과 결과 페이지 수를 안내하고 다운로드할 수 있습니다. PDF → HWP/HWPX는 원본 용지 크기와 추정 여백을 적용하며, 한컴에서 넘친 구역은 줄·문단 간격을 조정합니다. 결과의 페이지 수가 원본과 다르면 다운로드를 제공하지 않습니다. PDF·PPT·PPTX 결과는 페이지 수가 달라지면 다운로드를 제한합니다. 파일은 최대 32MB까지 한 번에 하나씩 처리하고, 업로드와 결과 임시 파일은 요청이 끝나면 정리합니다.

Word 문서를 PDF로 내보낼 때 표 셀보다 넓은 인라인 그림은 비율을 유지해 셀 안에 맞춥니다. 원본 파일에는 저장하지 않으며, 렌더링 비교에서는 이 레이아웃 변경도 차이로 표시합니다.

DOCX에서 HWP/HWPX로 변환할 때 Word에 저장된 자동 페이지 나눔을 원본 렌더링과 확인해, 다음 페이지의 제목이 앞 페이지로 당겨지지 않도록 합니다. 빈 페이지가 생길 경우 내용 없는 간격 문단을 줄인 결과도 비교합니다.

이 데모는 Windows PC의 루프백 주소에서 실행됩니다. 실제 변환과 비교에 로컬 한컴 및 Microsoft Office가 사용되므로 Cloudflare Pages나 Linux Oracle VM에 이 서버를 그대로 배포할 수 없습니다.

## 포함한 변환기

- HWP → DOCX는 [morningkim/hwp2docx 고정 커밋](https://github.com/morningkim/hwp2docx/tree/0daa40faacf8f025237724a45a81e042ff3c442b)의 스크립트를 사용합니다. 표와 압축 섹션 크기를 제한했습니다. 한글 5.0 바이너리의 본문과 표 중심이며 누락될 수 있는 서식과 개체가 있습니다.
- HWPX → DOCX는 [kangdacool HWPX 변환기 고정 커밋](https://github.com/kangdacool/hwpx-editing-skill/tree/07e50f43135d069aca0714b267f6e13e38718c5c)을 사용합니다. 필요한 XML 도우미를 분리하고, 단일 셀 표와 중첩 표도 Word 표로 남기도록 수정했습니다. 서식, 그림과 수식은 제한적으로 처리됩니다. 기존 내부 변환기는 CLI의 `--hwpx-structure`에서 사용할 수 있습니다.
- PDF → DOCX는 [ArtifexSoftware/pdf2docx](https://github.com/ArtifexSoftware/pdf2docx)에서 선이 있는 표와 배치로 추정되는 표를 모두 탐지합니다. 결과는 PDF 표의 구조에 따라 달라집니다. 위치를 우선하는 기존 방법은 CLI의 `--pdf-objects`에 남아 있습니다.
