"""Cycle a MODI+ LED on each new button press.

Sequence: red -> blue -> green -> red -> ...
The example uses a pressed-state rising edge, so holding the button does not
advance the color repeatedly. It is safe to stop with Ctrl+C.
"""

from __future__ import annotations

import argparse
import time
import traceback
from typing import Optional, Sequence


COLORS = (
    ("red", (255, 0, 0)),
    ("blue", (0, 0, 255)),
    ("green", (0, 255, 0)),
)


def module_label(module) -> str:
    module_type = getattr(module, "module_type", type(module).__name__)
    module_id = getattr(module, "id", None)
    if isinstance(module_id, int):
        return f"{module_type}(0x{module_id:X})"
    return str(module_type)


def find_first(modules, expected_type: str):
    for module in modules:
        if getattr(module, "module_type", "") == expected_type:
            return module
    raise RuntimeError(f"Missing {expected_type} module.")


def find_by_id_or_first(modules, expected_type: str, module_id: Optional[int]):
    if module_id is not None:
        for module in modules:
            if getattr(module, "id", None) == module_id:
                actual_type = getattr(module, "module_type", "")
                if actual_type != expected_type:
                    raise RuntimeError(
                        f"0x{module_id:X} is {actual_type!r}, expected {expected_type!r}."
                    )
                return module
        raise RuntimeError(f"Missing {expected_type}(0x{module_id:X}).")
    return find_first(modules, expected_type)


def run_self_test() -> int:
    expected = ("red", "blue", "green", "red", "blue", "green")
    actual = tuple(COLORS[index % len(COLORS)][0] for index in range(len(expected)))
    if actual != expected:
        raise AssertionError(f"Expected {expected}, got {actual}")
    print("Self-test passed: red -> blue -> green cycle.")
    return 0


def positive_int(value: str) -> int:
    try:
        return int(value, 0)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"invalid module ID: {value}") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--seconds",
        type=float,
        default=0.0,
        help="Stop after this many seconds; 0 means Ctrl+C stops the program.",
    )
    parser.add_argument(
        "--poll",
        type=float,
        default=0.05,
        help="Button polling interval in seconds (0.02 to 0.5).",
    )
    parser.add_argument("--button-id", type=positive_int, help="Optional button ID, e.g. 0x768.")
    parser.add_argument("--led-id", type=positive_int, help="Optional LED ID, e.g. 0xB9.")
    parser.add_argument("--self-test", action="store_true", help="Run without hardware.")
    return parser


def run_hardware(args: argparse.Namespace) -> int:
    import modi_plus

    bundle = None
    led = None
    try:
        bundle = modi_plus.MODIPlus(connection_type="serialport", verbose=False)
        time.sleep(2)
        modules = list(bundle.modules)
        print("Connected modules:")
        for module in modules:
            print(f"- {module_label(module)}")

        button = find_by_id_or_first(list(bundle.buttons), "button", args.button_id)
        led = find_by_id_or_first(list(bundle.leds), "led", args.led_id)
        button.prop_request_period = 0.1
        button.prop_samp_freq = 99
        led.turn_off()

        print(f"Using {module_label(button)} -> {module_label(led)}")
        print("Press and release the button: red -> blue -> green -> ...")
        print("Ctrl+C stops the program and turns the LED off.")

        color_index = -1
        was_pressed = bool(button.pressed)
        started_at = time.monotonic()
        while True:
            if args.seconds > 0 and time.monotonic() - started_at >= args.seconds:
                print("Timed run complete.")
                return 0

            pressed = bool(button.pressed)
            if pressed and not was_pressed:
                color_index = (color_index + 1) % len(COLORS)
                color_name, color = COLORS[color_index]
                led.rgb = color
                print(f"PRESS -> {color_name} {color}", flush=True)
            was_pressed = pressed
            time.sleep(args.poll)
    except KeyboardInterrupt:
        print("Stopped by user.")
        return 0
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        if led is not None:
            try:
                led.turn_off()
            except Exception:
                traceback.print_exc()
        if bundle is not None:
            try:
                bundle.close()
            except Exception:
                traceback.print_exc()


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.seconds < 0:
        raise SystemExit("--seconds cannot be negative")
    if not 0.02 <= args.poll <= 0.5:
        raise SystemExit("--poll must be between 0.02 and 0.5 seconds")
    return run_self_test() if args.self_test else run_hardware(args)


if __name__ == "__main__":
    raise SystemExit(main())
