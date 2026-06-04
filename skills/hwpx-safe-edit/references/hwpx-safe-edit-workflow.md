# HWPX Safe Edit Workflow

## 목적

원본 HWPX 양식을 보존하면서 필요한 텍스트만 바꾸는 절차를 고정한다.

## 절차

1. 원본 HWPX를 직접 수정하지 않는다.
2. `analyze_hwpx.py`로 ZIP 항목, 필수 파일, section XML, 표 수, 텍스트 노드 수를 확인한다.
3. `extract_text_map.py`로 `hp:t` 텍스트 노드 순서를 추출한다.
4. 바꿀 노드만 `replacements` JSON에 적는다.
5. `apply_text_map.py`로 새 HWPX를 만든다.
6. 기본값으로 `hp:linesegarray`를 제거한다.
7. `validate_hwpx.py`로 최종 파일을 검증한다.
8. 작업 폴더에 검증 로그를 남긴다.

## 원칙

- 구조 변경보다 텍스트 교체를 우선한다.
- 표 행/열 삭제는 사용자가 명시적으로 요구한 경우에만 한다.
- ZIP/XML 유효성만으로 최종 성공을 단정하지 않는다.
