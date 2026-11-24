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

CLI는 각 페이지의 크기와 144 DPI RGB 픽셀을 비교합니다. HWPX → DOCX는 페이지 수가 같으면 외형 차이가 있어도 저장합니다. 다른 경로의 기본 설정과 `--require-visual-match`는 용지 크기와 픽셀 일치까지 요구합니다. 일치 조건을 충족하지 못하면 종료 코드 2를 반환합니다. `--allow-differences`는 페이지 수가 같을 때 외형 차이를 허용하며, 편집 가능한 문서의 페이지 수도 달라질 수 있으면 `--allow-page-count-change`를 함께 사용합니다. PDF → HWP/HWPX는 원본과 같은 페이지 수를 요구합니다. 코드 1은 변환 또는 검증 실패입니다. 동일 판정은 실행 환경과 렌더링 설정에 한정됩니다.

HWPX → DOCX는 OWPML의 구역, 페이지 나눔, 문자 속성, 그림 참조, 표 및 셀 병합을 읽어 편집 가능한 Word 문서를 만듭니다. 각주·미주·수식이 있는 문서는 한컴의 Word 내보내기를 자동으로 사용하고, DOCX에 편집 가능한 개체가 남았는지 확인합니다. 원본에서 문서 끝에 배치한 미주가 Word에서 본문과 같은 쪽에 합쳐지면 새 구역을 시도하고, 원본 페이지 수가 복원된 결과를 채택합니다. `--hwpx-structure`를 사용하는 기존 명령도 지원합니다.

비교할 때 한컴 PDF가 원본 용지보다 조금 작게 내보낸 페이지는 문서에 기록된 용지 크기로 비교 영역을 복원합니다. 내보낸 크기는 보고서의 `source_render_size_pt`에 기록합니다. 각주·미주·수식이 있는 단일 구역 문서에서는 원본과 대응하는 본문 줄과 주석을 찾으면 글자 간격, 줄 위치와 주석 위치를 보정합니다. 단일 실선 각주 구분선은 원본에서 측정한 길이·두께·위치를 Word 구분선에 반영합니다. 수식 크기와 정렬은 HWPX 속성을 사용하며, 단순한 HancomEQN 수식의 변수는 Word 수식 안에서 기울임꼴을 복원합니다. 페이지 수와 편집 가능한 개체가 유지되며 픽셀 차이 합계가 줄어드는 후보를 저장합니다. 보고서의 `hwpx_rgb_error_before`와 페이지별 `rgb_error`로 차이를 비교할 수 있습니다.

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
