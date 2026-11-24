# FileConvert Skill

HWPX, HWP, Word, PDF, PowerPoint 문서를 로컬에서 양방향으로 변환하고 렌더링 결과를 비교하는 Codex 스킬입니다.

| 양방향 | 변환 엔진 |
| --- | --- |
| HWPX ↔ DOCX | 한컴 한글과 Word. HWPX → DOCX에 XML 구조 추출 옵션 제공 |
| HWP ↔ DOCX | 한컴 한글과 Word |
| PDF ↔ HWPX, PDF ↔ HWP | PDF의 내용과 배치를 한글 문서로 가져옵니다. 한글 문서에서 시작하면 PDF로 내보냅니다. |
| PDF ↔ DOCX, PDF ↔ DOC | PDF의 본문과 배치를 Word 문서로 재구성합니다. Word 문서에서 시작하면 PDF로 내보냅니다. |
| PDF ↔ PPTX, PDF ↔ PPT | PowerPoint PDF 내보내기. PDF에서 시작하면 페이지별 이미지를 슬라이드에 배치 |

Windows 데스크톱, Python 3.13(검증 버전)과 경로별로 설치·활성화된 한컴오피스 한글 및 Microsoft Word·PowerPoint가 필요합니다. PDF → PPTX 후보 생성과 HWPX 구조 추출은 Python으로 처리하지만 전체 CLI 검증에는 해당 오피스 앱이 필요합니다. Hancom 파일 접근 대화상자가 뜨는 환경에서는 한컴 공식 문서에 따라 등록한 경로 확인 모듈 이름을 `FILECONVERT_HANCOM_PATH_MODULE`에 지정할 수 있습니다.

```powershell
python -m pip install -r requirements.txt
python scripts/convert.py --list
python scripts/convert.py input.hwpx output.docx --report comparison.json
```

기본 모드는 각 페이지의 크기와 144 DPI RGB 픽셀을 비교합니다. 페이지 수나 크기 또는 픽셀이 하나라도 다르면 종료 코드 2를 반환하고 결과 문서 저장을 보류합니다. 비교 보고서에는 페이지별 크기와, 픽셀 크기가 같을 경우 달라진 영역의 경계가 기록됩니다. 페이지 수가 같을 때는 `--allow-differences`로 외형이 달라진 후보를 저장할 수 있습니다. 페이지 수가 달라지면 이 옵션을 사용해도 저장하지 않습니다. 코드 1은 변환 또는 검증 실패입니다. 동일 판정은 실행 환경과 렌더링 설정에 한정됩니다. 글꼴 파일과 편집 가능한 수식·각주·미주·필드의 보존은 대상 앱에서 별도로 확인해야 합니다.

HWPX에서 편집 가능한 본문을 먼저 얻고 싶으면 `--hwpx-structure`를 사용합니다. OWPML의 섹션, 문자 속성, 그림 참조, 표 및 셀 병합을 직접 읽습니다. 대표적인 글꼴·크기·색상과 용지 크기·여백을 매핑합니다. 모든 언어의 자간 값이 동일하면 Word 단위로 근사합니다. 언어별 자간·정확한 배치·수식·각주·미주·머리말·꼬리말은 이 경로에서 보존되지 않습니다. 이 옵션도 기본 비교를 거칩니다.

PDF를 Word나 한글 문서로 가져올 때 단, 머리말, 각주, 수식 및 줄바꿈이 재구성되는 방식은 원본과 다를 수 있습니다. 이미지로 된 PDF에 별도 OCR은 적용하지 않습니다. PowerPoint 경로는 페이지 이미지를 사용하므로 슬라이드의 텍스트는 개별 편집할 수 없습니다. 서로 다른 크기의 PDF 페이지는 첫 페이지 크기에 맞춰 여백을 추가할 수 있습니다.

Codex에 설치하려면 이 저장소를 사용자 스킬 폴더 아래 `file-convert-skill` 폴더로 복제하세요. 한컴과 Microsoft Office 자동화는 데스크톱 사용을 전제로 합니다.

참고한 방식은 HWP 바이너리와 HWPX XML 경로를 구분한 [morningkim/hwp2docx](https://github.com/morningkim/hwp2docx)와 HWPX의 문단·표를 읽어 DOCX를 구성한 [kangdacool/hwpx-editing-skill](https://github.com/kangdacool/hwpx-editing-skill/blob/main/skills/hwpx-editing/scripts/hwpx_to_docx.py)입니다. 이 저장소의 CLI 및 구조 추출 경로는 별도로 작성했습니다. 한컴 형식 처리는 [한컴 자동화 문서](https://developer.hancom.com/en-us/hwpautomation)를 참고했습니다.

## 검증 결과

Windows에서 합성 문서로 16개 방향 모두 결과 파일 생성을 확인했습니다. PDF로 내보내는 6개 방향은 해당 샘플의 렌더 비교를 통과했습니다. 나머지 10개 방향의 샘플은 외형 차이가 있어 엄격 모드 통과 대상으로 볼 수 없습니다.

별도의 Word → HWPX → Word 왕복 샘플에서 수식·각주·미주·머리말·꼬리말·페이지 번호의 문서 구조는 남았습니다. 자간 1.25pt는 1.20pt로 바뀌었고, 첫 한글 가져오기에서 인라인 그림이 누락됐습니다. 이 검증 결과는 문서마다 달라질 수 있습니다.
