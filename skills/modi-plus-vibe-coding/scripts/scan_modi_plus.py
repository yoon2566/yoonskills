"""Safely scan the modules visible through a MODI+ Network module."""

from __future__ import annotations

import argparse
import time
from typing import Iterable, Optional, Sequence


def module_label(module) -> str:
    module_type = getattr(module, "module_type", type(module).__name__)
    module_id = getattr(module, "id", None)
    if isinstance(module_id, int):
        return f"{module_type}(0x{module_id:X})"
    return str(module_type)


def run_self_test() -> int:
    class DemoModule:
        module_type = "led"
        id = 0xB9

    label = module_label(DemoModule())
    if label != "led(0xB9)":
        raise AssertionError(f"unexpected module label: {label}")
    print("Self-test passed: module label formatting.")
    return 0


def scan() -> int:
    import modi_plus

    bundle = None
    try:
        bundle = modi_plus.MODIPlus(connection_type="serialport", verbose=False)
        time.sleep(2)
        modules = list(bundle.modules)
        print(f"PyMODI+ {getattr(modi_plus, '__version__', 'unknown')}")
        print(f"Connected modules: {len(modules)}")
        for module in modules:
            print(f"- {module_label(module)}")
        return 0
    finally:
        if bundle is not None:
            bundle.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="Check formatting without opening a hardware connection.",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    return run_self_test() if args.self_test else scan()


if __name__ == "__main__":
    raise SystemExit(main())
