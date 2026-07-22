# 파이모디 바이브코딩 자료 분석

## 분석 범위와 방법

원본 폴더 `10_자료추가\파이모디_바이브코딩`에는 현재 PPTX가 없고 다음 21개 PDF가 있다.

- `파이모디_바이브코딩_1차시.pdf` ~ `20차시.pdf`: 20개 수업 덱, 총 186페이지
- `PyMODI+_라이브러리_레퍼런스.pdf`: 1개 API 덱, 21페이지
- 전체: 207페이지, 추출 문자 약 106,511자

PDF는 슬라이드 PDF 내보내기로 보고 `scripts/analyze_materials.py`로 페이지별 텍스트·문자 수·키워드·전체 텍스트를 추출했다. 각 PDF의 첫 두 페이지를 Poppler로 PNG 렌더링하고, 8차시와 레퍼런스의 API 페이지도 시각 확인했다. 일부 PDF 텍스트 추출에는 글꼴 때문에 `�`, 띄어쓰기 분리, 아이콘 문자가 생기므로 코드를 판정할 때는 렌더링된 페이지와 업스트림 소스를 함께 본다.

기계 생성 상세 보고서와 전체 추출 텍스트는 작업 폴더의 `15_modi_plus_skill_upgrade\analysis\materials\`에 있다. 다음에 PPTX가 추가되면 같은 스크립트가 PPTX XML도 분석한다.

## 20차시 학습 흐름

### 1차시 — AI란 무엇인가?

AI·머신러닝·딥러닝·생성형 AI의 관계, 일반 프로그래밍과 AI 프로그래밍의 차이, 학습과 추론, 실생활 AI 사례를 소개한다. ChatGPT·Claude·Gemini에 같은 질문을 던지고 답변의 길이·정확도·스타일을 비교한 뒤 간단한 Python 코드를 요청한다. 스킬에서는 “AI가 만든 코드도 실행·검증한다”는 원칙의 출발점으로 사용한다.

### 2차시 — AI 윤리: 잘 쓰는 법

편향, 자율주행·얼굴인식 같은 실패 사례, 개인정보·저작권, 환각, 딥페이크와 악용을 다룬다. O/X 퀴즈와 토론이 포함된다. 학생 프로젝트에서는 실제 이름·전화번호·얼굴 사진·API 키를 AI 프롬프트에 넣지 않고, AI 응답을 출처·실행 결과로 검증하도록 연결한다.

### 3차시 — 바이브코딩과 개발환경

자연어로 요구를 구체화하는 방법, 나쁜 프롬프트와 좋은 프롬프트 비교, Python·VS Code 설치와 Python 3.8.10 확인, Python 확장·폴더·첫 `print`를 다룬다. 교안은 PATH와 전역 설치를 안내하지만 이 스킬은 프로젝트 `.venv\Scripts\python.exe`를 명시해 재현성을 높인다.

### 4차시 — 파이썬 기초 ①

주석, 변수, 문자열·정수·실수·불리언 자료형, 산술·비교·논리 연산자, f-string, `input()`과 `int/float` 형변환을 배운다. 자기소개·계산기·문자열 카드·나이/출생연도 계산 실습으로 이어진다. 하드웨어 연결 전에 이 단계의 순수 Python을 통과시키는 것이 안전하다.

### 5차시 — 파이썬 기초 ②

`if/elif/else`, `while`, `for`, `break`, 리스트, 딕셔너리를 다룬다. 입장료·구구단·별 찍기·숫자 맞추기·점수 관리·학생 정보 카드 실습이 있다. MODI+에서는 센서 구간 분기, 색상 목록, 모듈별 매핑표에 그대로 연결된다.

### 6차시 — 파이썬 기초 ③

함수, 매개변수, `return`, `import`와 `random/time/math` 모듈을 배운다. 가위바위보와 할 일 관리 앱을 종합 프로젝트로 만들고 완료 체크·파일 저장 같은 바이브코딩 도전을 한다. 하드웨어 예제도 `set_color()`, `stop_all()` 같은 작은 함수로 나눌 근거가 된다.

### 7차시 — 파이모디 첫 연결

교안은 Setup(네트워크·배터리), Input(버튼·환경·IMU·ToF·다이얼·조이스틱), Output(LED·스피커·Display·모터)의 모듈 구성을 보여준다. Network를 USB로 연결하고 `MODIPlus()`, `bundle.modules`, 형식별 리스트, LED RGB, Display text를 처음 실행한다. 시각 검수에서 모듈 제품 그림과 코드 블록이 읽혔다. 자료에는 BLE 인자 `conn_type`이 보이지만 업스트림 생성자는 `connection_type`이므로 스킬에서는 후자를 사용한다.

### 8차시 — LED와 센서

Button의 `clicked`, `double_clicked`, `pressed`, `toggled`를 표로 비교하고 버튼·LED 기본 반응을 만든다. 핵심 실습은 버튼 클릭마다 빨강→초록→파랑→꺼짐으로 바뀌는 페이퍼 무드등과 밝기 3단계 무드등이다. 이어 Env의 빛·소리·온도·습도를 읽어 어두움·습도 조건에 따라 LED·Display를 바꾼다. 현재 예제는 학생 요구에 맞춰 빨강→파랑→초록 순환과 상승 에지·종료 정리를 제공한다.

### 9차시 — Display 활용하기

Display의 `.text`, `.reset()`, f-string과 소수점 포맷을 사용한다. 카운트다운·전광판, Env 온도·습도 실시간 온도계, ToF 거리별 위험 표시, Dial 값에 따른 LED 밝기·색 선택·게이지를 만든다. 공통 패턴은 `while → 센서값 읽기 → Display 출력 → sleep`이며, 초기값과 갱신 주기를 기록한다.

### 10차시 — 다중 센서 조합

`and/or/not`로 빛·소리·온도·습도를 함께 판단하고, 조건 우선순위 매핑표를 먼저 작성한다. 예시는 어둡고 시끄러운 긴급 경보, 고온, 어두움, 소음 상태를 LED·Display로 구분하는 스마트 경보 시스템이다. AI에게 코드를 요청할 때도 센서→조건→출력 표를 먼저 전달하고, 조건이 겹칠 때 `if/elif` 순서를 명시한다.

### 11차시 — 무선통신으로 파이모디 연결

USB 유선과 Bluetooth·Wi-Fi 개념을 비교하고 Network의 BLE 모드, 컴퓨터 Bluetooth 설정, `MODIPlus(connection_type="ble", network_uuid=...)`, 원격 LED를 다룬다. 외부 전원·페어링·UUID를 별도 확인해야 하며, 업스트림은 Windows Python 3.12+ BLE 호환성 제한을 문서화한다. 초급 학생은 USB 성공 후 무선으로 올라간다.

### 12차시 — 키보드로 자동차 운전하기

Motor의 속도·각도와 양수/음수 방향, 두 모터의 전진·회전 패턴, `pynput` 키보드 이벤트를 다룬다. 교안 예제는 모터를 바로 움직일 수 있으므로 실제 수업에서는 차체를 공중에 띄우고 저속·짧은 시간·수동 정지를 먼저 적용한다. ID와 배열 순서만으로 좌우를 추정하지 않는다.

### 13차시 — 인터넷에서 날씨 가져오기

API와 크롤링 차이, OpenWeatherMap 가입·API 키, `requests`, JSON 파싱을 다룬다. API 키를 코드에 공개하지 않고 환경변수나 학생 로컬 설정으로 분리한다. 네트워크가 없는 순수 JSON 샘플 테스트와 MODI+ 연결 테스트를 나누면 오류 원인을 좁힐 수 있다.

### 14차시 — 날씨로 LED·Display 제어

Clear·Clouds·Rain·Snow 같은 날씨 상태를 노랑·회색·파랑·흰색 RGB와 Display 메시지로 매핑하고 자동 갱신·깜빡임을 통합한다. 먼저 고정된 가짜 API 응답으로 분기와 색 매핑을 테스트한 뒤 실제 API와 하드웨어를 연결한다. 키·인터넷·LED·Display를 한 번에 디버깅하지 않는다.

### 15차시 — 목소리로 명령하기

Speech-to-Text 개념, `speech_recognition`, `pyaudio`, 마이크 입력과 인식 실패 처리를 다룬다. 마이크 권한·소음·인터넷 서비스 의존성을 먼저 순수 음성 예제로 확인하고, 아직 MODI+를 연결하지 않은 상태에서도 인식 문자열 로그를 남긴다.

### 16차시 — 음성으로 파이모디 제어

“켜줘/꺼줘/빨간색/파란색” 같은 음성 명령을 LED·Display 동작으로 매핑한다. 음성→텍스트→명령 분기→하드웨어 출력의 네 단계를 로그로 분리하고, 인식 실패·알 수 없는 명령·Ctrl+C에서 LED를 끈다.

### 17차시 — 카메라로 보는 로봇

OpenCV로 웹캠을 열고 프레임을 읽어 평균 밝기·움직임을 계산하며, 어두움·보통·밝음에 따라 LED를 바꾼다. 카메라 번호·권한·창 닫기·`numpy` 같은 외부 조건을 먼저 검증하고, MODI+는 밝기 임계값을 만든 뒤 연결한다.

### 18차시 — 포즈·제스처 인식

MediaPipe Hands·Pose·Face Mesh 개념, 21개 손 관절, 손가락 수·주먹·V 제스처 판별, 제스처→LED·Display 제어를 다룬다. 카메라 프레임·AI 판정·MODI+ 출력의 책임을 함수로 나누고, 인식되지 않을 때 안전한 기본 출력(off)을 사용한다.

### 19차시 — 노벨 엔지니어링 ①

이야기 읽기→문제 발견→해결 설계→프로토타입의 4단계를 사용한다. 헨젤과 그레텔의 거리 경보, 빨간 모자의 이미지 인식, 성냥팔이 소녀의 온도 알림, 라푼젤의 버튼·모터 승강기, 잭과 콩나무의 IMU 안전모 같은 예시가 있다. HMW(How Might We) 질문으로 “누가·어디서·무엇 때문에·어떤 어려움”인지 구체화한다.

### 20차시 — 노벨 엔지니어링 ②

HMW 질문을 `센서 → 조건 → 출력` 매핑표로 바꾸고, AI에게 정확한 입력 모듈·출력 모듈·우선순위·종료 조건을 전달해 코드를 설계한다. 학생 워크시트, 프롬프트 템플릿, 1~20차시 회고가 포함된다. 최종 작품은 코드보다 문제 정의·매핑표·테스트 기록·수정 이유까지 제출한다.

## PyMODI+ 레퍼런스 PDF 분석

21페이지 레퍼런스는 소개·입력 모듈·출력 모듈·조합 예제로 구성된다.

- 소개: 공식 저장소, Python 3.7+, Windows/macOS/Linux, Serial 기본 연결과 BLE, `pip install pymodi-plus`, `python -m modi_plus --help/--tutorial/--inspect/--usage`
- 공통: `bundle = modi_plus.MODIPlus()`, `bundle.modules`, `bundle.leds/dials/speakers...`, 첫 모듈 `[0]`
- 입력: Button 4상태, Dial `turn/speed`, Env `temperature/humidity/illuminance/volume`, IMU 각도·각속도·가속도·진동, ToF `distance`, Joystick `x/y/direction`
- 출력: LED RGB, Display text/reset/write/draw, Speaker tune/frequency/volume/music, Motor angle/speed/stop
- 조합: Dial→LED, ToF→Speaker, 센서값→출력의 반복 루프

레퍼런스와 공식 0.4.2 소스를 대조한 차이는 다음과 같다.

1. 레퍼런스의 BLE 예제는 `conn_type`으로 보이지만 소스 생성자 인자는 `connection_type`이다.
2. 레퍼런스 LED 표는 개별 `red/green/blue`를 0~100으로, `rgb`를 0~255로 적는다. 초급 코드는 RGB 튜플을 사용한다.
3. 레퍼런스 Motor 표의 `torque`는 현재 업스트림 `Motor` 클래스에 없다. 실제 소스의 `set_angle`, `set_speed`, `append_angle`, `stop`만 사용한다.
4. Env 온도 범위 표가 문서마다 다르므로 임계값은 실측하고, 버전·모듈 펌웨어에 맞춰 기록한다.

공식 소스와 설치 버전의 함수·속성 목록은 [pymodi-plus-upstream.md](pymodi-plus-upstream.md)에 고정해 두었다.
