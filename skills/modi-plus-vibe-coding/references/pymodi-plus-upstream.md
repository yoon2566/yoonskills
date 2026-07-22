# PyMODI+ 공식 저장소 대조 결과

## 조사 provenance

- 저장소: `https://github.com/LUXROBO/pymodi-plus`
- 조사한 branch: `master`
- 조사한 commit: `c9f169dc6a2836b50a2dde45fa9bc92183afc9e8`
- package metadata: `pymodi-plus 0.4.2`
- 로컬 clone: `15_modi_plus_skill_upgrade/source/pymodi-plus`
- 조사일: 2026-07-22

공식 소스는 MIT 라이선스의 Python API다. 이 문서는 소스를 복사해 배포하는 문서가 아니라 학생용 스킬이 어떤 공개 API와 안전 규칙을 기준으로 하는지 고정하는 대조표다. 업스트림이 바뀌면 commit과 설치 버전을 다시 기록한다.

## 설치와 연결

`setup.py`는 Python `>=3.7`을 선언한다. 현재 `requirements.txt`의 핵심은 다음과 같다.

- 공통: `pyserial==3.5`, `nest-asyncio==1.5.4`, `websocket-client==1.2.3`, `packaging>=21.3`
- Windows: `bleak==0.13.0`, `winusbcdc==1.4`
- macOS: `bleak==0.13.0`
- Linux: `pexpect`

Windows 학생은 우선 Python 3.8.10 + 프로젝트 `.venv` + USB Network 연결을 사용한다. 업스트림 troubleshooting 문서는 Python 3.12+ Windows BLE에서 `bleak-winrt` 호환 문제가 있어 BLE 관련 테스트를 제한한다고 설명한다.

실제 생성자는 다음과 같다.

~~~python
import modi_plus

bundle = modi_plus.MODIPlus(connection_type="serialport", verbose=False)
~~~

소스의 `SerialportTask`는 `list_modi_ports()`를 호출한다. Windows에서 `connection_util.py`는 일반 serial 포트의 MODI+ VID/PID(`0x2FDE`, `0x0003`) 또는 `MODI+ Network Module` 설명과 WinUSB 경로를 함께 검사한다. 따라서 장치 관리자에 COM 번호가 보이지 않는 것만으로 연결 실패라고 단정하지 않는다.

BLE는 소스상 `connection_type="ble", network_uuid="..."`이다. 일부 README·교안 예제의 `conn_type`은 소스와 일치하지 않으므로 학생 코드에는 사용하지 않는다.

`MODIPlus.close()`는 실행 스레드와 연결을 닫는다. 모든 라이브 하드웨어 프로그램은 `finally`에서 `bundle.close()`를 호출한다.

## 모듈 API 대조표

### 공통 접근

`MODIPlus`는 `modules`, `networks`, `batterys`, `envs`, `imus`, `buttons`, `dials`, `joysticks`, `tofs`, `displays`, `motors`, `leds`, `speakers` 목록 속성을 제공한다. 단일 ID가 필요하면 `bundle.button(id)`, `bundle.led(id)`, `bundle.motor(id)`처럼 형식 검사를 하는 메서드를 사용한다. 배열 `[0]`은 현재 연결 순서의 첫 모듈이지 차체의 영구 좌우 위치가 아니다.

### 입력 모듈

| 모듈 | 실제 소스 속성 | 스킬에서의 기본 사용 |
|---|---|---|
| Button | `clicked`, `double_clicked`, `pressed`, `toggled` | `pressed` 상승 에지 또는 클릭 이벤트 |
| Dial | `turn`, `speed` | 0~100 값, LED 밝기/색 선택 |
| Env | `illuminance`, `temperature`, `humidity`, `volume` | 값 읽기 후 임계값·Display 출력 |
| Env RGB(v0.4.x, 하드웨어 조건부) | `red`, `green`, `blue`, `white`, `black`, `color_class`, `brightness`, `rgb`, `raw_red/green/blue/white`, `raw_rgb`, `set_rgb_mode` | 먼저 지원 여부와 펌웨어를 확인 |
| IMU | `angle_x/y/z`, `angle`, `angular_vel_x/y/z`, `angular_velocity`, `acceleration_x/y/z`, `acceleration`, `vibration` | 센서값 단독 출력 후 알림 |
| ToF | `distance` | 거리 구간별 경보 |
| Joystick | `x`, `y`, `direction` | deadzone·속도 상한 후 자동차 제어 |

### 출력 모듈

| 모듈 | 실제 소스 속성/메서드 | 안전·정확성 메모 |
|---|---|---|
| LED | `rgb`, `red`, `green`, `blue`, `set_rgb`, `turn_on`, `turn_off` | 초급 코드는 `rgb=(R,G,B)`와 `turn_off()` 사용 |
| Display | `text`, `write_text`, `write_variable_xy`, `write_variable_line`, `draw_picture`, `draw_dot`, `set_offset`, `reset` | 화면 지우기를 정상 종료에도 호출 |
| Speaker | `tune`, `frequency`, `volume`, `set_tune`, `play_music`, `stop_music`, `pause_music`, `resume_music`, `reset` | 소리 실험은 작은 volume으로 제한 |
| Motor | `angle`, `target_angle`, `speed`, `target_speed`, `set_angle`, `set_speed`, `append_angle`, `stop` | `finally`에서 `stop()`과 `set_speed(0)` |

## 중요한 구현 세부

### Button

소스의 Button property는 상태 property 번호 2를 읽어 `clicked`, `double_clicked`, `pressed`, `toggled`를 각각 2바이트 값으로 해석한다. `pressed`는 누르고 있는 상태를 읽는 데 적합하고, 계속 누르는 동안 한 번만 반응해야 하면 이전 상태와 비교한다. `prop_request_period`와 `prop_samp_freq`는 빠른 입력 실험에서 조정할 수 있지만 기본값을 먼저 사용한다.

### LED

소스의 `rgb` setter는 `set_rgb(red, green, blue)`를 호출한다. 공식 examples와 교안은 RGB 예제에 0~255를 사용한다. `turn_on()`은 소스에서 `(100, 100, 100)`을 보내며, `turn_off()`는 `(0, 0, 0)`을 보낸다. 개별 채널 표가 0~100으로 보이는 문서가 있으므로 초급 프로젝트는 튜플 한 가지 표현으로 통일한다.

### Motor

`set_angle(target_angle, target_speed=70)`의 source guard는 angle 0~360, speed 0~100이다. `set_speed(target_speed)`는 signed integer를 보내므로 음수·양수 방향은 차체 장착 방향에 따라 실제로 관찰해야 한다. `append_angle`과 `stop`도 제공한다. 공식 레퍼런스의 `torque` 표기는 현재 0.4.2 source와 맞지 않아 사용하지 않는다.

모터는 입력이 중립이면 즉시 0으로 보내고, 버튼·예외·통신 종료·Ctrl+C에서 모든 모터를 정리한다. 공중 테스트의 성공은 바닥 주행 성공이 아니다.

## 소스 파일 근거

- `modi_plus/modi_plus.py`: 생성자, connection type, module lists, `close`
- `modi_plus/module/module.py`: module ID/type, property cache, `ModuleList`
- `modi_plus/module/input_module/button.py`: Button 4상태
- `modi_plus/module/input_module/joystick.py`: x/y/direction
- `modi_plus/module/input_module/env.py`: 일반·RGB Environment API
- `modi_plus/module/output_module/led.py`: RGB와 off/on
- `modi_plus/module/output_module/motor.py`: angle/speed/append/stop
- `modi_plus/task/serialport_task.py`: USB/serial receive/send lifecycle
- `modi_plus/util/connection_util.py`: MODI+ port와 WinUSB 탐색
- `examples/basic_usage_examples/`: 공식 버튼·LED·조이스틱·모터 예제
- `docs/troubleshooting/PYTHON_313_FIX.md`, `WINDOWS_BLE_FIX.md`: 최신 Python/Windows 주의점

## 학생에게 적용하는 결론

1. 업스트림 API 이름을 기준으로 코드를 만들되, 설치된 버전과 연결 모듈을 먼저 출력한다.
2. 교안의 아이디어는 유지하되 외부 API·BLE·카메라·음성은 하드웨어와 분리된 작은 테스트를 먼저 만든다.
3. 문서에 없는 속성(예: `motor.torque`)을 AI가 생성하면 실행 전에 소스·`hasattr`·공식 examples로 확인한다.
4. `bundle.modules` 목록, LED 동작, 버튼 이벤트, 모터 방향을 각각 별도 증거로 기록한다.
