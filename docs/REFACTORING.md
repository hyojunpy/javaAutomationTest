# 변환 결과를 유지하는 구조 개선

브랜치: `refactor/conversion-ui`

## 유지한 규칙

- 일반형과 시간연장형의 메뉴 해석, 휴일 처리, 생일 식단, 주별 시트 분리
- 유형·연령별 기존 파일명 인식 규칙과 연월 추론 기본값
- 신규/기존 템플릿 선택 순서, 재료 검색, 수량 계산, 수식과 서식
- 공개 정적 변환 API 및 JSON 파일/문자열 입력 지원

## 코드 구조

- `ConversionEngine`: JavaFX, Swing, CLI의 공통 HWP → JSON → XLSX 흐름
- `SourceMetadata`: 유형·연령 판별과 기본 출력 파일명. 일반형/시간연장형의 기존 연령 판별 차이를 의도적으로 유지
- `TemplateCatalog`: 템플릿 이름과 home/new → home/old → cwd/new → cwd/old 탐색 순서
- `WorkbookResources`: 로딩 스트림을 닫고 Workbook은 호출부의 try-with-resources로 관리
- 파서와 Excel 작성기: 공개 API 호출마다 별도 인스턴스를 만들어 문서, 연월, 스타일 캐시를 작업 단위로 관리
- 기존 `app.home` 설정은 외부 호출과 CLI 호환을 위해 읽기만 지원. GUI와 엔진은 경로를 직접 전달
- 식단 파서의 콘솔 디버그 출력을 SLF4J DEBUG로 변경. 오류는 사용자 작업 내역과 로거에 기록

## 검증

`mvn test`는 파일명/템플릿 탐색/입력 오류/파서 작업 상태 분리를 검사합니다.
`mvn -Pregression verify`는 예시 8개를 실제 변환하고 원본 결과와 비교합니다. 일반형은 수 분 걸릴 수 있습니다.
기준 파일은 `src/test/resources/golden`에 있으며 상세 비교 범위와 생성 원칙은 그 폴더의 README를 참고하세요.
예시 외 문서의 정확성을 보장하는 테스트는 아니며, 원본 결과의 업무적 정확성을 검증하는 테스트와도 구분됩니다.

## UI

유형·연령과 파일 경로 확인, 경과 시간, 작업 단계, 완료 후 Excel/폴더 열기, 오류 시 작업 내역 자동 펼치기를 제공합니다.
변환 중 파일 교체 및 드래그 앤 드롭을 차단합니다. 기존 파일 덮어쓰기는 사용자에게 확인합니다.
스크롤 레이아웃으로 작은 창에서도 버튼과 작업 내역에 접근할 수 있습니다.

## EXE 빌드

Windows x64 JDK 21, Maven, JavaFX 21.0.4 JMODs를 준비하고 프로젝트 루트에서 실행합니다.

```powershell
.\scripts\build-windows.ps1 -JdkHome 'C:\tools\jdk-21' -MavenHome 'C:\tools\apache-maven-3.9.9' -JavaFxJmods 'C:\tools\javafx-jmods-21.0.4'
```

`-RunRegression`을 추가하면 전체 회귀 테스트를 통과한 뒤 패키징합니다.
`-OutputDirectory`를 생략하면 `portable/<생성된 ID>`를 사용합니다. 기존 배포 폴더를 덮어쓰지 않습니다.
결과는 실행용 EXE 폴더와 ZIP이며 설치 마법사가 있는 설치용 EXE는 기존 README 절차를 사용합니다.

현재 변경은 구조와 사용성 개선이며, 변환 속도 최적화는 별도 프로파일링과 동일 결과 검증 후 진행합니다.
