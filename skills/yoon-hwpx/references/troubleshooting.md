# Troubleshooting

## 한글에서 손상/변조 경고가 뜬다

가능성이 높은 원인:

- 본문 텍스트는 바뀌었지만 기존 `hp:linesegarray`가 남아 있다.
- ZIP 첫 항목 `mimetype`이 바뀌었다.
- section XML은 well-formed지만 표 구조나 내부 참조가 깨졌다.

우선 확인:

```powershell
.\.venv\Scripts\python.exe .\scripts\validate_hwpx.py .\output\result.hwpx --expect-no-linesegarray
```

## 한글이 깨져 보인다

PowerShell 기본 인코딩 문제일 수 있다.

```powershell
$env:PYTHONUTF8='1'
$env:PYTHONIOENCODING='utf-8'
```

## XML 검증은 통과하지만 결과가 이상하다

XML 구조 검증은 문서 레이아웃 보존을 보장하지 않는다. 원본과 다음 값을 비교한다.

- section 파일 목록
- `hp:t` 텍스트 노드 수
- `hp:tbl` 표 수
- `hp:linesegarray` 잔여 수
- 바뀐 텍스트와 남아 있으면 안 되는 텍스트
