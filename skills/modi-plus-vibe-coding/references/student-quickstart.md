# 학생용 시작 안내

이 스킬은 공개 저장소 `https://github.com/yoon2566/yoonskills`의 `skills/modi-plus-vibe-coding`에 배포한다. 저장소를 clone한 뒤 프로젝트 루트에 Python 3.8.10 가상환경을 만들면 같은 명령을 재현할 수 있다.

## 1. clone과 가상환경

PowerShell에서 실행한다.

```powershell
git clone https://github.com/yoon2566/yoonskills.git
Set-Location .\yoonskills

# Python 3.8.10이 설치되어 있을 때
py -3.8 -m venv .venv
& .\.venv\Scripts\python.exe --version
```

`py -3.8`을 찾지 못하면 Python 3.8.10 64비트를 설치한 뒤 새 PowerShell을 연다. 전역 Python으로 실행하지 말고 아래처럼 항상 프로젝트 실행기를 지정한다.

```powershell
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install pymodi-plus==0.4.2
& .\.venv\Scripts\python.exe -c "import modi_plus; print(modi_plus.__version__)"
```

## 2. 스킬 폴더 사용

Codex가 clone한 저장소를 작업 폴더로 열면 `skills/modi-plus-vibe-coding/SKILL.md`를 읽을 수 있다. 전역 Codex 스킬 폴더에 복사해야 하는 환경에서는 원본을 보존할 수 있도록 먼저 대상이 없는지 확인한다.

```powershell
$repoRoot = (Get-Location).Path
$codexSkills = Join-Path $env:USERPROFILE '.codex\skills'
$skillTarget = Join-Path $codexSkills 'modi-plus-vibe-coding'
New-Item -ItemType Directory -Path $codexSkills -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repoRoot 'skills\modi-plus-vibe-coding') -Destination $skillTarget -Recurse -Force
```

## 3. 하드웨어 없이 먼저 확인

```powershell
& .\.venv\Scripts\python.exe .\skills\modi-plus-vibe-coding\scripts\scan_modi_plus.py --self-test
& .\.venv\Scripts\python.exe .\examples\button_led_rgb_cycle.py --self-test
& .\.venv\Scripts\python.exe -m py_compile .\examples\button_led_rgb_cycle.py
```

예상되는 핵심 출력은 다음과 같다.

```text
Self-test passed: module label formatting.
Self-test passed: red -> blue -> green cycle.
```

## 4. MODI+ 연결과 첫 작품

Network 모듈에 USB를 연결하고 Button과 LED를 같은 MODI+ 체인에 붙인다. 먼저 모터나 자동차를 연결하지 않은 정적인 작품으로 시작한다.

```powershell
& .\.venv\Scripts\python.exe .\skills\modi-plus-vibe-coding\scripts\scan_modi_plus.py
```

스캔 결과에 `button(...)`과 `led(...)`가 보이면 직접 실행한다.

```powershell
& .\.venv\Scripts\python.exe .\examples\button_led_rgb_cycle.py --seconds 60
```

또는 [examples\run_button_led_rgb_cycle.bat](../../examples/run_button_led_rgb_cycle.bat)을 더블클릭한다. BAT는 현재 저장소의 `.venv\Scripts\python.exe`를 사용한다. 실행 중인 콘솔에서 버튼을 눌렀다 떼면 `PRESS -> red`, `PRESS -> blue`, `PRESS -> green`이 순서대로 출력되어야 한다. 끝낼 때는 `Ctrl+C`를 누른다.

ID가 여러 개일 때는 스캔 결과의 ID를 명시한다.

```powershell
& .\.venv\Scripts\python.exe .\examples\button_led_rgb_cycle.py --button-id 0x768 --led-id 0xB9 --seconds 60
```

## 5. 학생 기록 양식

각 실습 폴더의 README에 다음 네 가지를 남긴다.

1. 실행한 명령과 Python·`pymodi-plus` 버전
2. 실제 감지된 모듈 종류와 ID
3. 콘솔 이벤트 로그와 눈으로 관찰한 LED/Display 결과
4. 아직 누르거나 움직여 보지 못한 입력, 아직 바닥에서 시험하지 않은 자동차 등 미검증 항목

버튼 이벤트가 없으면 바로 새 작품을 만들지 말고 스캔 → `pressed` 값 출력 → polling 간격 확인 순서로 진단한다. raw property는 고장 진단용으로만 사용한다.

## 6. 자동차로 확장할 때

모터가 있는 학생은 [전체 작업 이력](work-history.md)과 [공식 API 대조](pymodi-plus-upstream.md)를 먼저 읽는다. 차체를 바닥에 놓기 전에 바퀴를 공중에 띄우고, 모터 ID별 양수·음수 방향을 기록한다. `stop()`, `set_speed(0)`, 중립·버튼·`Ctrl+C`·통신 예외 정지가 없는 자동차 코드는 실행하지 않는다.
