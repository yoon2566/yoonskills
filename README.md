# Yoon Codex Skills

윤정환의 공개 Codex 스킬 허브입니다. 반복해서 사용할 수 있는 스킬은 이 저장소에서 관리하고, 대형 스크립트 묶음이나 독립 배포가 필요한 스킬은 별도 저장소로 유지합니다.

## 빠른 선택

| 작업 | 권장 스킬 | 상태 |
|---|---|---|
| 여러 교육 도구와 산출물을 함께 조정 | `skills/edukit` | 사용 권장 |
| 기존 Canva 교육 템플릿의 내용만 교체 | `skills/canva-education-template-edit` | 도구 연결 시 사용 |
| HWPX 새 문서 생성·분석·검증 | `skills/hwpx` | 사용 권장 |
| 기존 HWPX 양식 보존, 이미지 삽입, 한컴 손상 복구 | [`yoon_hwpx_3`](https://github.com/yoon2566/yoon_hwpx_3) | **현재 대표판** |
| 강원SW미래채움 사진 교체 반복 업무 | `skills/swhwpx` | 특수 업무 전용 |
| 가벼운 HWPX 텍스트 안전 편집 | `skills/yoon-hwpx` | 레거시 호환용 |

## 이 저장소의 스킬

### `edukit`

교육 과정·문서·프레젠테이션·Google Workspace·3D·영상 등 두 개 이상의 작업 영역을 조정하는 상위 오케스트레이션 스킬입니다. 단일 작업에는 전문 스킬을 우선합니다.

### `canva-education-template-edit`

기존 Canva 템플릿의 디자인을 유지하면서 교육용 텍스트를 교체하는 스킬입니다. Canva 편집 도구가 연결된 환경에서 사용합니다.

### `hwpx`

HWPX 문서 생성, 템플릿 분석, 텍스트 추출, 패키징, 검증을 위한 범용 스킬입니다. 원본 없이 새 문서를 만드는 작업은 이 스킬을 우선합니다.

### `swhwpx`

강원SW미래채움 서명록·활동사진 교체처럼 특정 반복 업무를 위한 특수 스킬입니다. 범용 HWPX 작업에는 사용하지 않습니다.

### `yoon-hwpx`

초기 HWPX 원본 보존형 안전 편집 스킬입니다. 새 작업은 기능이 확장된 `yoon_hwpx_3`을 우선하고, 기존 호출 호환이 필요할 때만 사용합니다.

## 공개 저장소 분류

| 저장소 | 분류 | 권장 처리 |
|---|---|---|
| [`yoonskills`](https://github.com/yoon2566/yoonskills) | 공개 스킬 허브 | 계속 사용 |
| [`yoon_hwpx_3`](https://github.com/yoon2566/yoon_hwpx_3) | 최신 독립 HWPX 스킬 | 대표판으로 사용 |
| [`yoon_hwpx_2`](https://github.com/yoon2566/yoon_hwpx_2) | HWPX 2세대 | 레거시·비교용으로 유지 후 추후 보관 처리 검토 |
| [`yoon_hwpx`](https://github.com/yoon2566/yoon_hwpx) | HWPX 초기 안전 편집 키트 | 레거시·학습 기록용 |
| [`swmc-codex-practice1`](https://github.com/yoon2566/swmc-codex-practice1) | Codex 교육 실습 자료 | 스킬 저장소와 분리 유지 |
| `google-drive-plugin-copy-report` | 플러그인 실험 보고서 | 스킬이 아닌 연구 기록으로 분류 |

기계별 원격 제어, IP 주소, 장치 ID, 개인 경로가 포함된 스킬은 공개 허브에 넣지 않고 비공개 저장소로 분리합니다.

## 설치

### 이 저장소의 스킬 하나 설치

```powershell
git clone https://github.com/yoon2566/yoonskills.git
New-Item -ItemType Directory -Force "$HOME\.agents\skills" | Out-Null
Copy-Item -Recurse ".\yoonskills\skills\edukit" "$HOME\.agents\skills\edukit"
```

`edukit` 대신 설치할 스킬 폴더명을 지정합니다.

### 최신 HWPX 스킬 설치

```powershell
git clone https://github.com/yoon2566/yoon_hwpx_3.git "$HOME\.agents\skills\yoon-hwpx-3"
```

설치 후 새 Codex 세션에서 해당 스킬이 인식되는지 확인합니다.

## 관리 원칙

1. 한 스킬은 하나의 독립 폴더로 관리합니다.
2. 각 스킬은 최소한 `SKILL.md`를 포함합니다.
3. 복잡한 실행 로직은 `scripts/`, 참고 자료는 `references/`, UI 메타데이터는 `agents/openai.yaml`에 둡니다.
4. 공개 스킬에는 이메일, 장치 이름, 사설 IP, 원격 접속 ID, API 키, 개인 절대 경로를 넣지 않습니다.
5. 샘플 문서와 이미지는 공개 전에 개인정보와 기관 정보를 제거합니다.
6. 새 버전이 이전 버전을 대체하면 이 README에서 `현재 대표판`과 `레거시`를 명확하게 표시합니다.
7. 특정 기관·수업·컴퓨터에만 쓰는 스킬은 범용 스킬과 분리합니다.

## 후속 정리 대상

- `한글문서테스트원본/`의 이미지와 HWPX가 공개용 익명 샘플인지 재확인합니다.
- `yoon_hwpx`와 `yoon_hwpx_2`는 README에 최신판 이동 안내를 추가한 뒤 GitHub Archive 사용을 검토합니다.
- `skills/yoon-hwpx`와 `yoon_hwpx_3`의 중복 기능은 최신판 기준으로 점차 정리합니다.
- 플러그인 실험 보고서는 향후 `research/` 전용 저장소로 모으는 방안을 검토합니다.
