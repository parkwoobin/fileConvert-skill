# FileConvert Skill

[![Agent Skill](https://img.shields.io/badge/Agent_Skill-SKILL.md-8A2BE2)](SKILL.md)
[![Works with Claude Code and Codex](https://img.shields.io/badge/Works_with-Claude_Code_%7C_Codex-orange)](#claude-code와-codex에-설치)
[![Python 3.13](https://img.shields.io/badge/Python-3.13-3776AB)](#claude-code와-codex에-설치)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D4)](#claude-code와-codex에-설치)

HWPX, HWP, Word, PDF, PowerPoint 문서를 로컬에서 양방향으로 변환하고 렌더링 결과를 비교하는 Claude Code와 Codex용 문서 변환 스킬입니다.

| 변환 경로 | 변환 방식 |
| --- | --- |
| HWPX ↔ DOCX | HWPX XML의 본문과 표를 Word 문서로 구성합니다. DOCX에서 시작하면 한글 문서로 저장합니다. |
| HWP ↔ DOCX | HWP 5.0의 본문과 표를 Word 문서로 구성합니다. DOCX에서 시작하면 한글 문서로 저장합니다. |
| PDF ↔ HWPX, PDF ↔ HWP | PDF의 내용과 배치를 한글 문서로 가져옵니다. 한글 문서에서 시작하면 PDF로 내보냅니다. |
| PDF ↔ DOCX, PDF ↔ DOC | PDF의 본문과 배치를 Word 문서로 재구성합니다. Word 문서에서 시작하면 PDF로 내보냅니다. |
| PDF ↔ PPTX, PDF ↔ PPT | PowerPoint PDF 내보내기. PDF에서 시작하면 페이지별 이미지를 슬라이드에 배치 |
| DOC → DOCX | Microsoft Word에서 DOCX로 저장 |
| HWP → HWPX | 한컴 한글에서 HWPX로 저장 |

Windows 데스크톱, Python 3.13(검증 버전)과 경로별로 설치·활성화된 한컴오피스 한글 및 Microsoft Word·PowerPoint가 필요합니다. PDF → PPTX 후보 생성과 HWPX 구조 추출은 Python으로 처리하지만 전체 CLI 검증에는 해당 오피스 앱이 필요합니다. Hancom 파일 접근 대화상자가 뜨는 환경에서는 한컴 공식 문서에 따라 등록한 경로 확인 모듈 이름을 `FILECONVERT_HANCOM_PATH_MODULE`에 지정할 수 있습니다.

```powershell
python -m pip install -r requirements.txt
python scripts/convert.py --list
python scripts/convert.py input.hwpx output.docx --report comparison.json
```

기본 모드는 각 페이지의 크기와 144 DPI RGB 픽셀을 비교합니다. 페이지 수나 크기 또는 픽셀이 하나라도 다르면 종료 코드 2를 반환하고 결과 문서 저장을 보류합니다. 비교 보고서에는 페이지별 크기와, 픽셀 크기가 같을 경우 달라진 영역의 경계가 기록됩니다. 페이지 수가 같을 때는 `--allow-differences`로 외형이 달라진 후보를 저장할 수 있습니다. 페이지 수가 달라진 편집용 문서를 받으려면 `--allow-page-count-change`도 추가하세요. PDF에서 HWP/HWPX로 변환할 때는 이 옵션을 사용해도 원본과 같은 페이지 수를 요구합니다. 코드 1은 변환 또는 검증 실패입니다. 동일 판정은 실행 환경과 렌더링 설정에 한정됩니다. 글꼴 파일과 편집 가능한 수식·각주·미주·필드의 보존은 대상 앱에서 별도로 확인해야 합니다.

HWPX에서 편집 가능한 본문을 먼저 얻고 싶으면 `--hwpx-structure`를 사용합니다. OWPML의 섹션, 문자 속성, 그림 참조, 표 및 셀 병합을 직접 읽습니다. 대표적인 글꼴·크기·색상과 용지 크기·여백을 매핑합니다. 모든 언어의 자간 값이 동일하면 Word 단위로 근사합니다. 언어별 자간·정확한 배치·수식·각주·미주·머리말·꼬리말은 이 경로에서 보존되지 않습니다. 이 옵션도 기본 비교를 거칩니다.

PDF를 Word나 한글 문서로 가져올 때 단, 머리말, 각주, 수식 및 줄바꿈이 재구성되는 방식은 원본과 다를 수 있습니다. 이미지로 된 PDF에 별도 OCR은 적용하지 않습니다. PowerPoint 경로는 페이지 이미지를 사용하므로 슬라이드의 텍스트는 개별 편집할 수 없습니다. 서로 다른 크기의 PDF 페이지는 첫 페이지 크기에 맞춰 여백을 추가할 수 있습니다.

## Claude Code와 Codex에 설치

스킬 이름은 `file-convert-skill`이며, 실행 절차는 [SKILL.md](SKILL.md)에 있습니다. 스킬 작성과 형식 검증에 Codex의 `skill-creator`를 사용했습니다.

이 저장소의 `SKILL.md`, `scripts/`, `requirements.txt`를 한 폴더에 둡니다. 설치 폴더 이름은 `file-convert-skill`로 지정합니다.

| 환경 | 사용자 스킬 경로 | 직접 호출 |
| --- | --- | --- |
| Claude Code | `~/.claude/skills/file-convert-skill/` | `/file-convert-skill` |
| Codex | `~/.agents/skills/file-convert-skill/` | `$file-convert-skill` |

Windows PowerShell에서는 사용할 앱의 경로 하나를 선택해 복제합니다.

```powershell
$skillRoot = Join-Path $env:USERPROFILE '.agents/skills' # Codex
# $skillRoot = Join-Path $env:USERPROFILE '.claude/skills' # Claude Code
New-Item -ItemType Directory -Force -Path $skillRoot | Out-Null
git clone https://github.com/parkwoobin/fileConvert-skill.git (Join-Path $skillRoot 'file-convert-skill')
```

필요한 변환 프로그램과 Python 의존성을 설치한 뒤 새 세션에서 스킬을 호출합니다. 한컴과 Microsoft Office 의존 경로는 해당 앱을 실행할 수 있는 Windows 환경이 필요합니다.

[Claude Code 스킬 문서](https://code.claude.com/docs/en/skills) · [Codex 설치 경로 안내](https://developers.openai.com/api/reference/cli)

## 검증 결과

Windows에서 처음 구현한 변환 방향 16개의 샘플로 결과 파일 생성을 확인했습니다. 이 가운데 PDF로 내보내는 6개 방향은 렌더 비교를 통과했고, 나머지 10개 방향은 외형 차이가 있었습니다. 이후 추가한 DOC → DOCX와 HWP → HWPX는 본문과 표가 있는 1페이지 샘플에서 렌더 비교를 통과했습니다.

별도의 Word → HWPX → Word 왕복 샘플에서 수식·각주·미주·머리말·꼬리말·페이지 번호의 문서 구조는 남았습니다. 자간 1.25pt는 1.20pt로 바뀌었고, 첫 한글 가져오기에서 인라인 그림이 누락됐습니다. 이 검증 결과는 문서마다 달라질 수 있습니다.

## PDF의 글자와 도형을 Word 개체로 분리

고정 배치 PDF에는 `python scripts/convert.py input.pdf output.docx --pdf-objects --allow-differences --report report.json`을 사용할 수 있습니다. 화면 위치에 맞춰 수정 가능한 텍스트 상자, Word 도형, 개별 그림을 배치합니다. 같은 흐름의 여러 줄은 한 텍스트 상자로 묶고, 원본 렌더링을 기준으로 숨은 텍스트와 근접 중복을 걸러냅니다. 자간과 폰트는 근사치이며 해당 글꼴이 설치되어 있어야 원본에 가깝게 표시됩니다. 일부 PDF 그림은 잘라낸 이미지로 들어갑니다. 클리핑, 겹침, 그룹의 복합 효과는 부분적으로만 처리하고 PDF의 표와 문단 구조를 재현하지 않습니다. 개체 편집 검증은 Windows용 Microsoft Word에서 했습니다.
