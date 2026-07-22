# Yoon Skills

This public repository contains reusable Codex skills. The MODI+ skill is a
Windows/Python teaching workflow for LUXROBO MODI+ hardware.

## MODI+ quick start

~~~powershell
git clone https://github.com/yoon2566/yoonskills.git
Set-Location .\yoonskills
py -3.8 -m venv .venv
& .\.venv\Scripts\python.exe -m pip install pymodi-plus==0.4.2
& .\.venv\Scripts\python.exe .\skills\modi-plus-vibe-coding\scripts\scan_modi_plus.py --self-test
& .\.venv\Scripts\python.exe .\examples\button_led_rgb_cycle.py --self-test
~~~

Read [the MODI+ skill](skills/modi-plus-vibe-coding/SKILL.md) before connecting
hardware. The student example is
[examples/button_led_rgb_cycle.py](examples/button_led_rgb_cycle.py), and the
full material report is [reports/materials_analysis.md](reports/materials_analysis.md).

The skill uses the official `LUXROBO/pymodi-plus` API, Python 3.8.10, project
`.venv`, real module scans, hardware-free self-tests, and explicit motor
safety gates. The current material survey covers 20 session PDFs plus one
PyMODI+ reference PDF; the report explains the PDF/PPTX distinction.
