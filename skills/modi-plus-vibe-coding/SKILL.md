---
name: modi-plus-vibe-coding
description: Build, debug, teach, and document MODI+ physical-computing projects with Python and Codex. Use when connecting MODI+ modules, discovering hardware, controlling LEDs, buttons, displays, sensors, joysticks, speakers, or motors, creating always-on programs, calibrating a MODI+ car, or turning an experiment into a safe student project.
---

# MODI+ 바이브 코딩

MODI+ 작품은 `아이디어 → 작은 실험 → 실제 관찰 → 수정 → 작품` 순서로 만든다. 하드웨어를 추측하지 말고 연결 목록과 실행 결과를 확인한다. 학생에게는 결과만 주지 말고, 무엇을 확인했고 무엇이 아직 물리적으로 검증되지 않았는지 함께 기록한다.

## 이 스킬의 사용 범위

- Windows PowerShell과 프로젝트 `.venv`를 기준으로 한다.
- 공식 `LUXROBO/pymodi-plus` 소스와 20차시 교안의 학습 순서를 함께 사용한다.
- 버튼·LED 같은 정적인 첫 작품은 빠르게 증명하되, 모터·자동차는 별도 안전 절차를 통과한 뒤에만 실행한다.
- 학생이 clone한 저장소에서 바로 재현할 수 있도록 명령, 파일 위치, self-test, 종료 방법을 남긴다.

다음 참고 자료는 필요할 때만 읽는다. 먼저 이 파일의 흐름을 적용하고, API·수업 순서·과거 판단의 근거가 필요할 때 해당 문서를 연다.

- [학생용 시작 안내](references/student-quickstart.md)
- [20차시 자료 분석](references/teaching-materials.md)
- [공식 저장소 대조 결과](references/pymodi-plus-upstream.md)
- [MODI+ API 최소 패턴](references/modi-plus-api.md)
- [전체 작업 이력과 검증 상태](references/work-history.md)

## 1. 시작 전에 환경을 고정한다

PowerShell에서 프로젝트 루트를 기준으로 실행한다.

```powershell
& .\.venv\Scripts\python.exe --version
& .\.venv\Scripts\python.exe -c "import modi_plus; print(modi_plus.__version__); print(modi_plus.__file__)"
git status --short --branch
rg --files
```

- 전역 Python, 전역 pip, `pip install --user`를 사용하지 않는다.
- 이 프로젝트의 검증 기준은 Python 3.8.10 64비트와 `pymodi-plus 0.4.2`다. 다른 버전이면 먼저 실제 설치 버전과 의존성을 기록한다.
- `pymodi-plus`를 새로 설치할 때는 프로젝트 환경 안에서만 실행한다.

```powershell
& .\.venv\Scripts\python.exe -m pip install pymodi-plus==0.4.2
```

## 2. 교안 순서로 난이도를 올린다

학생의 작업을 아래 순서에 맞춘다. 매 단계의 코드는 하드웨어 없이도 가능한 부분부터 self-test한다.

1. 1~3차시: AI 개념, AI 윤리, 좋은 프롬프트, Python·VS Code 환경 확인
2. 4~6차시: 주석·변수·자료형·연산자·문자열·입력, 조건·반복·리스트·딕셔너리, 함수·모듈·작은 종합 프로젝트
3. 7차시: Network를 USB로 연결하고 모듈 종류와 ID를 실제로 스캔한 뒤 LED·Display 하나를 출력 증명
4. 8차시: Button의 `clicked`, `double_clicked`, `pressed`, `toggled`를 비교하고 LED 색 순환 작품으로 확장
5. 9~10차시: Display·Env·ToF·Dial을 붙이고 `센서 → 조건 → 출력` 매핑표와 우선순위 로직을 만든다.
6. 11~18차시: BLE, 모터·키보드, 날씨 API, 음성, 카메라, MediaPipe를 한 번에 섞지 말고 외부 의존성과 하드웨어를 분리해 단계별로 검증한다.
7. 19~20차시: 이야기 속 문제를 HMW 질문으로 바꾸고, 매핑표·프롬프트·프로토타입·회고까지 문서화한다.

세부 페이지 수, 각 차시의 실습, 자료의 API 오류 가능성은 [교안 분석](references/teaching-materials.md)을 따른다.

## 3. 실제 연결을 먼저 증명한다

사용자의 아이디어를 먼저 한 문장으로 적는다.

> 어떤 입력이 들어오면, 어떤 출력이 발생하고, 언제 멈추는가?

그 다음 최소 스캔을 실행한다. `bundle.modules`의 배열 위치를 영구적인 하드웨어 의미로 사용하지 말고 `module_type`과 `id`를 함께 출력한다.

```python
import time
import modi_plus

bundle = None
try:
    bundle = modi_plus.MODIPlus(connection_type="serialport", verbose=False)
    time.sleep(2)
    for module in list(bundle.modules):
        print(f"{module.module_type}(0x{module.id:X})")
finally:
    if bundle is not None:
        bundle.close()
```

재사용 스크립트는 [scripts/scan_modi_plus.py](scripts/scan_modi_plus.py)다.

```powershell
& .\.venv\Scripts\python.exe .\skills\modi-plus-vibe-coding\scripts\scan_modi_plus.py
& .\.venv\Scripts\python.exe .\skills\modi-plus-vibe-coding\scripts\scan_modi_plus.py --self-test
```

연결이 실패하면 COM 포트만 보지 않는다. 공식 Windows 구현은 USB VID/PID와 WinUSB 경로도 검사하므로 Network 전원, 체인 접점, USB 드라이버, Python 환경, 다른 프로세스의 연결 점유 순서로 확인한다.

## 4. 출력 하나부터 관찰한다

LED는 다음 두 표현 중 `rgb` 또는 `set_rgb` 하나로 통일한다. 현재 설치 버전과 공식 소스에서 RGB 튜플/인자 예제는 0~255를 사용한다.

```python
led.rgb = (255, 0, 0)
led.set_rgb(0, 255, 0)
led.turn_off()
```

전원만으로 LED가 켜진 것과 Python 명령으로 색이 바뀐 것은 다르다. 실행 로그와 눈으로 본 색 순서를 함께 기록한다. `turn_on()`은 소스상 각 채널을 100으로 설정하는 레거시 편의 메서드이므로 “255 최대 밝기”와 같은 의미로 단정하지 않는다.

Display는 `display.text = "..."`와 `display.reset()`을 사용한다. 센서값은 f-string으로 출력하고 갱신 간격을 둔다.

## 5. 입력 이벤트는 중복을 막아 연결한다

`pressed`는 누르고 있는 동안의 상태이고, `clicked`·`double_clicked`·`toggled`는 버튼 이벤트 확인에 사용한다. 버튼을 계속 누르고 있어도 한 번만 색을 바꾸려면 `pressed`의 상승 에지를 저장한다.

```python
was_pressed = False
color_index = -1
colors = ((255, 0, 0), (0, 0, 255), (0, 255, 0))

while True:
    pressed = bool(button.pressed)
    if pressed and not was_pressed:
        color_index = (color_index + 1) % len(colors)
        led.rgb = colors[color_index]
    was_pressed = pressed
    time.sleep(0.05)
```

학생이 바로 실행할 수 있는 완성 예제는 [examples/button_led_rgb_cycle.py](../../examples/button_led_rgb_cycle.py)다. 동작은 `빨강 → 파랑 → 초록 → 빨강 → ...`이고 `Ctrl+C`, 시간 제한, 예외에서 LED를 끈다.

## 6. 기능을 작게 확장하고 증거를 남긴다

다음 순서를 지킨다.

1. 연결 스캔: 실제 모듈의 종류·ID·개수를 기록한다.
2. 출력 단독 테스트: LED 한 색 또는 Display 한 문장을 보낸다.
3. 입력 단독 테스트: 버튼·다이얼·센서값을 콘솔에 출력한다.
4. 입력→출력: 한 조건과 한 출력만 연결한다.
5. 제한 시간·로그·종료 조건을 추가한다.
6. 하드웨어 없는 self-test와 `py_compile`을 실행한다.
7. 라이브 실행 결과와 아직 관찰하지 못한 항목을 README/report에 적는다.

AI가 만든 코드는 그대로 믿지 않는다. 실행 전 API 이름, 인자 범위, 연결 모듈, 종료 경로를 소스와 설치 버전으로 대조한다. 교안과 소스가 다르면 소스·실행 결과를 우선하고 차이를 문서화한다.

## 7. 모터와 자동차는 별도 안전 게이트를 둔다

- 모터가 보이면 곧바로 돌리지 말고 ID와 개수를 먼저 확인한다.
- 처음에는 차체를 들어 바퀴를 완전히 공중에 띄우고 속도 8, 0.4초 정도의 짧은 테스트만 한다.
- 실제 양수·음수 회전 방향과 좌·우 모터 위치를 관찰해 `direction_notes.md`에 적기 전에는 바닥 주행을 허용하지 않는다.
- 조이스틱 deadzone, 속도 상한, 중립 정지, 버튼 정지, `Ctrl+C`, 통신 예외 정지를 모두 구현한다.
- 모든 모터 명령은 `try/finally`에서 `stop()`과 `set_speed(0)`으로 정리한다.
- 방향이 확인되지 않은 모터의 ID를 배열 위치만 보고 왼쪽·오른쪽으로 단정하지 않는다.

기존 자동차 코드의 안전 플래그 패턴은 `--run --confirm-airborne`, 이후 별도 검증을 거친 `--floor-run --confirm-floor`다. 물리적 바닥 주행을 했다는 증거가 없으면 성공으로 보고하지 않는다.

## 8. 항상 실행되는 프로그램의 종료 규칙

긴 실행 프로그램은 기본적으로 시간 제한 실행을 먼저 제공한다. 백그라운드 실행은 직접 실행과 로그 확인이 끝난 뒤에만 추가한다.

```python
bundle = None
try:
    bundle = modi_plus.MODIPlus(connection_type="serialport")
    # 입력을 읽고 출력한다.
except KeyboardInterrupt:
    print("Stopped by user.")
finally:
    if bundle is not None:
        for motor in list(bundle.motors):
            try:
                motor.stop()
                motor.set_speed(0)
            except Exception:
                pass
        for led in list(bundle.leds):
            try:
                led.turn_off()
            except Exception:
                pass
        bundle.close()
```

## 9. 자주 생기는 불일치와 문제 해결

- 공식 README와 일부 교안은 BLE 인자 이름을 `conn_type`으로 적지만 현재 소스의 생성자는 `connection_type`이다.
- 라이브러리 레퍼런스는 `motor.torque`를 표에 넣지만 현재 0.4.2 `Motor` 소스에는 해당 속성이 없다.
- 레퍼런스의 LED 개별 채널 범위(0~100)와 RGB 튜플 예시(0~255)가 섞여 있다. 초급 코드는 `led.rgb = (R, G, B)`로 통일한다.
- Environment RGB 센서는 하드웨어 버전에 따라 지원 여부가 다르므로 `red`, `green`, `blue`, `raw_rgb`를 무조건 호출하지 않는다.
- Python 3.12+ Windows BLE는 업스트림 troubleshooting 문서에서 제한 사항이 있으므로 이 작업 흐름은 Python 3.8.10과 USB 연결을 기본으로 한다.
- 버튼이 반응하지 않으면 프로세스가 살아 있는지, `bundle.buttons`에 버튼이 있는지, `pressed`가 0↔100으로 바뀌는지, polling 간격이 있는지 순서대로 확인한다. raw property 접근은 진단용으로만 사용한다.

## 10. 완료 판정

작품을 “완료”라고 말하기 전에 다음을 모두 확인한다.

- 실행한 Python 경로와 `pymodi-plus` 버전을 기록했다.
- 실제 연결 모듈의 종류와 ID를 기록했다.
- 하드웨어 없는 self-test 또는 순수 로직 테스트가 통과했다.
- 라이브 실행에서는 명령 로그와 관찰된 출력이 일치했다.
- 정상 종료·`Ctrl+C`·예외에서 LED와 모터가 안전 상태가 됐다.
- 물리적으로 확인하지 못한 항목을 성공으로 포장하지 않았다.
- 학생이 clone 후 사용할 PowerShell 명령과 더블클릭용 BAT 위치가 README에 있다.

전체 이력의 완료/미완료 상태는 [references/work-history.md](references/work-history.md)에 유지한다.
