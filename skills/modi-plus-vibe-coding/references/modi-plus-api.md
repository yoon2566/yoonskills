# MODI+ API 최소 패턴

이 문서는 `pymodi-plus 0.4.2`와 공식 소스를 대조한 초급용 참고다. 실행 전에 프로젝트 `.venv`의 Python 경로, 설치 버전, 실제 모듈 목록을 확인한다. 전체 근거는 [pymodi-plus-upstream.md](pymodi-plus-upstream.md)에 있다.

## 연결과 스캔

~~~python
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
~~~

`bundle.buttons[0]`처럼 목록의 첫 항목을 쓰기 전에 목록을 출력한다. 여러 같은 종류가 있거나 차체 방향이 중요하면 ID와 `module_type`을 함께 검사한다.

~~~python
def find_module(modules, module_id, expected_type):
    for module in modules:
        if getattr(module, "id", None) == module_id:
            actual_type = getattr(module, "module_type", "")
            if actual_type != expected_type:
                raise RuntimeError(
                    f"0x{module_id:X} is {actual_type!r}, expected {expected_type!r}"
                )
            return module
    raise RuntimeError(f"Missing {expected_type}(0x{module_id:X})")
~~~

## LED

초급 프로젝트는 RGB 튜플을 사용한다. 공식 소스의 `rgb` setter와 `set_rgb`가 같은 출력 경로를 사용한다.

~~~python
led.rgb = (255, 0, 0)
led.set_rgb(0, 255, 0)
led.turn_off()
~~~

`turn_on()`은 현재 source에서 `(100, 100, 100)`을 보내는 레거시 메서드다. 문서의 개별 채널 0~100 표와 RGB 0~255 예제가 섞여 있으므로 둘을 같은 범위라고 설명하지 않는다.

## Button

`pressed`는 누르는 동안의 상태다. 한 번 누를 때만 동작시키려면 상승 에지를 사용한다.

~~~python
was_pressed = False
while True:
    pressed = bool(button.pressed)
    if pressed and not was_pressed:
        print("new press")
    if not pressed and was_pressed:
        print("release")
    was_pressed = pressed
    time.sleep(0.05)
~~~

소스가 제공하는 이벤트 속성은 `clicked`, `double_clicked`, `pressed`, `toggled`다. `pressed`가 계속 0이면 먼저 프로세스, 연결된 Button 목록, polling 간격, raw 진단 순서를 확인한다.

## Display와 센서

~~~python
display.text = f"T:{env.temperature:.1f}C"
time.sleep(1)
display.reset()

distance = tof.distance
if distance < 10:
    led.rgb = (255, 0, 0)
else:
    led.turn_off()
~~~

센서값을 `while True` 안에서 읽을 때는 `time.sleep()`으로 갱신 주기를 제한한다. Environment RGB 속성은 하드웨어·펌웨어 조건부이므로 지원 여부를 확인한 뒤 사용한다.

## Joystick과 Motor

Joystick 값은 `x`, `y`가 대략 -100~100이고 `direction`은 `up/down/left/right/origin`이다. 자동차에서는 deadzone과 속도 상한을 먼저 적용한다.

~~~python
try:
    left_motor.set_speed(left_speed)
    right_motor.set_speed(right_speed)
finally:
    left_motor.stop()
    right_motor.stop()
    left_motor.set_speed(0)
    right_motor.set_speed(0)
~~~

Motor 소스의 `set_angle` 범위는 각도 0~360, 속도 0~100이다. `set_speed`의 음수·양수 방향은 차체 장착 방향에 따라 다르므로 공중에서 관찰한다. 현재 소스에는 `motor.torque`가 없다.

## 진단용 raw property

아래처럼 `_get_property`를 직접 읽는 것은 일반 작품 코드가 아니라 Button 입력이 들어오는지 확인하는 진단용이다.

~~~python
import struct

raw = bytes(button._get_property(2)[:8]).ljust(8, b"\x00")
clicked, double_clicked, pressed, toggled = struct.unpack("HHHH", raw)
print(clicked, double_clicked, pressed, toggled)
~~~

일반 코드에서는 공개 속성(`button.pressed` 등)을 사용하고, 진단 결과에는 raw 값과 공개 속성의 차이를 적는다.

## 종료 정리

~~~python
try:
    # hardware work
    pass
finally:
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
~~~
