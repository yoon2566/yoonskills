# MODI+ 작업 이력과 현재 검증 상태

이 문서는 `C:\Users\user\Desktop\모디`에서 번호를 붙여 진행한 작업을 스킬의 설계 근거로 요약한 것이다. “코드가 있다”와 “실제 하드웨어 동작을 관찰했다”를 구분한다.

## 번호별 이력

| 폴더 | 작업 | 확인된 결과 | 남은 주의점 |
|---|---|---|---|
| `1_modi_codex_vibe_start` | MODI 작품을 아이디어·입력·출력·안전 조건으로 쪼개는 시작 구조 | AI에게 줄 프로젝트 프롬프트, 모듈 목록, 시뮬레이션 우선 원칙을 만들었다. | 실제 모듈 연결 전에는 어댑터와 하드웨어를 추측하지 않는다. |
| `2_modi_plus_connection_check` | Network 연결과 Python 환경 확인 | MODI+ Network가 Python에서 인식되고 `pymodi-plus 0.4.2` import가 됐다. 당시 Python 3.14에서 BLE 의존성은 USB 점검에서 제외했다. | COM 포트가 없어도 Windows WinUSB 경로일 수 있으므로 `list_modi_ports()`와 실제 초기화를 함께 확인한다. |
| `3_modi_plus_led_check` | 첫 LED 점등 | Network·LED를 인식하고 LED 색상 테스트 스크립트가 정상 종료했다. | 단순 전원 점등과 Python 제어를 분리해서 판정한다. |
| `4_modi_plus_control_proof` | LED 제어 증명 | Python 명령 순서와 관찰 순서를 대조하는 보고서와 로그 구조를 만들었다. | 실행 로그만으로 빛을 봤다고 주장하지 않는다. |
| `5_modi_plus_button_led` | 버튼→LED 첫 작품 | Network·Button·LED 스캔, 누름/해제 반응 코드, 시간 제한 실행 구조를 만들었다. | 자동 실행 로그에는 버튼 이벤트가 없었던 기록이 있어 실제 눌림 관찰은 별도 확인이 필요했다. |
| `6_modi_plus_button_diagnosis` | 버튼 raw 상태 진단 | 버튼을 누르면 `pressed_value=100`, 놓으면 `0`으로 돌아가 Python까지 입력이 전달됨을 확인했다. | raw API는 일반 작품 코드가 아니라 고장 진단용으로만 쓴다. |
| `7_modi_plus_always_on` | 계속 실행되는 버튼·LED 프로그램 | 터미널 실행, 백그라운드 시작·중지, 로그·PID·정리 구조를 확인했다. | 항상 켜두기 전 직접 실행과 안전한 종료를 먼저 증명한다. |
| `8_modi_plus_car_scan` | MODI+ 자동차 형태와 모듈 스캔 | 사진 기준 차량 조립과 두 모터 후보를 확인했다. | 왼쪽/오른쪽 위치와 양수/음수 실제 방향은 아직 단정하지 않았다. |
| `9_modi_plus_motor_calibration` | 모터 방향 캘리브레이션 | 모터 `0xB1B`, `0x746`를 대상으로 dry-run과 속도 8·0.4초 공중 테스트 명령을 만들었다. | 실제 방향 관찰과 `direction_notes.md` 확정, 바닥 주행은 아직 남아 있다. |
| `10_자료추가` | 파이모디 바이브코딩 수업자료 수집 | 현재 폴더의 20차시 PDF와 PyMODI+ 레퍼런스를 원본 보존 상태로 분석했다. | 현재 정확한 대상은 PPTX가 아니라 21개 PDF다. |
| `11_modi_new_chat_handoff` | 새 대화 인수인계 | Python 3.8.10, `.venv`, 설치 버전, 자동차 미완료 상태, 안전 명령을 다음 작업자가 이어갈 수 있게 기록했다. | 하드웨어 연결 snapshot은 시점마다 달라지므로 매번 다시 스캔한다. |
| `12_modi_plus_joystick_car` | 조이스틱 차동 자동차 코드 설계 | `joystick.x/y`, 버튼 정지, 모터 `set_speed/stop`, deadzone·속도 상한·예외 정리와 self-test를 넣었다. | 실제 모터 방향과 바닥 주행은 물리적으로 검증되지 않았다. |
| `13_modi_plus_button_rgb_cycle` | 버튼 RGB 순환 | Python 3.8.10 64비트, `pymodi-plus 0.4.2`, Network `0xB22`, Button `0x768`, LED `0xB9` 연결과 30초 정상 종료·LED off를 확인했다. | 해당 실행 동안 눌림 로그가 없어 색상 전환 자체의 물리적 관찰은 아직 미확정이다. |
| `14_modi_vibe_practice` | 교안 초반 Python 연습 | `01_hello_world.py`를 프로젝트 `.venv`에서 compile/run하고 Python 3.8.10 실행을 확인했다. | 4~6차시 문법 실습을 이어서 하드웨어 작업과 분리한다. |
| `15_modi_plus_skill_upgrade` | 학생용 스킬 업그레이드 | 공식 저장소 clone, 21개 PDF·207페이지 분석, 재사용 스캔·버튼 RGB 예제, 교안/API/이력 참고문서를 만들고 공개 skill 저장소 배포를 준비 중이다. | 최종 push 후 clone·self-test·스킬 validator를 다시 확인한다. |

## 시점별 하드웨어 snapshot

서로 다른 날짜의 스캔을 합쳐 현재 상태라고 말하지 않는다.

- 자동차 초기 snapshot: Display `0x66`, Network `0xB22`, Button `0x768`, Joystick `0x244`, LED `0xB9`, Motor `0xB1B`, Motor `0x746`
- 이후 연결 점검: Network `0xB22`, Speaker `0x631`, Button `0x768`, LED `0xB9`, ToF `0xC44`
- 2026-07-22 새 스캔으로 확인한 10모듈 snapshot: Network `0xB22`, Speaker `0x631`, Env `0x340`, Display `0x66`, Dial `0x4AC`, Button `0x768`, Joystick `0x244`, LED `0xB9`, IMU `0xF56`, ToF `0xC44`; 모터는 보이지 않았다.

따라서 자동차 코드는 모터가 다시 보일 때까지 실행하지 않고, 정적인 Button·LED 예제는 최신 스캔에서 해당 모듈이 보일 때만 실행한다.

## 이력에서 일반화한 규칙

1. 설치 버전·Python 경로·모듈 ID를 모든 보고서에 적는다.
2. 출력 단독 → 입력 단독 → 입력·출력 결합 순으로 확장한다.
3. 버튼은 상태를 이벤트로 바꿀 때 상승 에지를 사용하고 polling 간격을 둔다.
4. 모터는 방향을 관찰하기 전까지 바닥에 놓지 않는다.
5. 모든 하드웨어 출력은 정상 종료와 예외에서 neutral/off 정리를 실행한다.
6. 결과가 없었던 실행도 실패가 아니라 다음 진단을 정하는 증거로 기록한다.
