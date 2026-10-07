# 조리지시서 작성 자동화 프로그램

어린이집 식단 HWP 파일을 읽고, 기존 조리지시서 Excel 양식을 활용해 XLSX 파일을 만드는 Windows 데스크톱 프로그램입니다. 일반형과 시간연장형, 만1–2세와 만3–5세 식단의 반복적인 조리지시서 작성 작업을 돕습니다.

## 주요 기능

- **HWP → XLSX 변환**: 식단을 분석하고 조리지시서 템플릿을 바탕으로 Excel 파일을 생성합니다.
- **유형 자동 선택**: 입력 파일명에 `시간연장`이 있으면 시간연장형, 없으면 일반형으로 처리합니다.
- **연령별 양식 적용**: 파일명의 연령 정보를 이용해 만1–2세 또는 만3–5세 양식을 선택합니다.
- **파일 선택 및 드래그 앤 드롭**: JavaFX 화면에서 HWP 파일을 선택하거나 끌어다 놓을 수 있습니다.
- **저장 위치 지정**: 출력 경로 변경과 변환 후 폴더 열기를 지원합니다.
- **진행 상태 확인**: 변환 상태와 경로 정보를 화면과 로그에서 확인할 수 있습니다.

## 사용 방법

1. 프로그램을 실행합니다.
2. **파일 선택**을 누르거나 HWP 파일을 화면에 끌어다 놓습니다.
3. **출력 XLSX** 경로를 확인합니다. 필요하면 **변경**으로 저장 위치를 지정합니다.
4. **변환 시작**을 누릅니다.
5. **변환 완료**가 표시되면 생성된 Excel 파일을 열어 내용을 확인합니다.

처음 선택한 HWP 파일과 같은 폴더에 `<원본 파일명>_수정.xlsx`가 기본 출력 경로로 제안됩니다. 한 번에 HWP 파일 하나를 변환합니다.

### 입력 파일명

유형·연령·연월 판단에 파일명을 사용하므로, 저장소 예시와 같은 이름을 유지하세요.

```text
2026년 2월 일반형(만1-2세).hwp
2026년 2월 일반형(만3-5세).hwp
2026년 2월 시간연장형(만1-2세).hwp
2026년 2월 시간연장형(만3-5세).hwp
```

입력은 `.hwp` 형식이며 `.hwpx`는 지원하지 않습니다. 저장소의 식단·조리지시서 양식에 맞춰 구현되어 있으므로, 구조가 다른 HWP 문서는 파서나 템플릿 수정이 필요할 수 있습니다.

## 개발 환경에서 실행

현재 Maven 설정은 Windows용 JavaFX를 사용합니다. 아래 명령은 **Windows PowerShell**에서 실행합니다.

필요한 도구:

- JDK 21: 프로젝트의 JavaFX 21.0.4 구성에 맞춰 사용합니다. Java 소스의 컴파일 대상은 `release 17`로 설정되어 있습니다.
- Maven: `mvn` 명령을 사용할 수 있도록 PATH를 설정합니다.
- Git: 저장소를 복제할 때 사용합니다.

```powershell
git clone https://github.com/hyojunpy/javaAutomationTest.git
cd javaAutomationTest
java -version
mvn -version
mvn clean compile
if ($LASTEXITCODE -ne 0) { throw "컴파일 실패" }
Copy-Item -LiteralPath .\input -Destination .\target\input -Recurse -Force
mvn javafx:run
```

**`target/input` 복사가 필요합니다.** 개발 실행 시 클래스가 `target/classes`에 생성되고, `AppHomeResolver`는 클래스/JAR 위치와 그 바로 위 폴더에서 `input`을 찾습니다. 저장소 루트의 `input`만 있으면 GUI 변환에서 `input 폴더가 없습니다` 오류가 발생할 수 있습니다. `mvn clean`을 실행한 뒤에는 다시 복사하세요.

### JAR 빌드

```powershell
mvn clean package
```

빌드가 성공하면 의존성을 포함한 `target/menu-hwp-to-excel-1.0.0-all.jar`가 생성됩니다. JavaFX 실행 환경과 외부 `input` 폴더도 필요하므로, 이 JAR 하나만 복사하는 것으로 배포가 완료되지는 않습니다.

## Windows EXE 만들기

`mvn javafx:run`은 프로그램을 실행하고, `mvn clean package`는 JAR를 만듭니다. **EXE를 만들려면 아래 `jpackage` 단계까지 실행해야 합니다.**

| 결과물 | 용도 | 추가 도구 |
| --- | --- | --- |
| 실행용 EXE 폴더 (`app-image`) | 압축을 풀고 바로 실행 | JDK 21, JavaFX JMODs |
| 설치용 EXE (`exe`) | 설치 마법사로 프로그램 설치 | 위 도구 + WiX Toolset 3.x |

두 방식 모두 Java 런타임을 포함하므로, 완성된 배포본을 사용하는 PC에 별도로 Java를 설치할 필요는 없습니다. Windows용 패키지는 Windows에서 생성하세요.

### 1. 패키징 준비

- `JAVA_HOME`을 **JDK 21 설치 폴더**로 설정하고 `mvn -version`에서도 같은 JDK를 사용하는지 확인합니다.
- Windows x64용 **JavaFX 21.0.4 JMODs**를 준비합니다. `pom.xml`의 JavaFX 버전 및 JDK의 아키텍처와 맞추세요. SDK의 `lib` 폴더가 아니라 `.jmod` 파일이 들어 있는 폴더가 필요합니다.
- 아래 예시의 `$fxJmods`를 실제 JMODs 폴더로 수정합니다.
- 프로젝트 루트에서 같은 PowerShell 세션으로 순서대로 실행합니다.

JavaFX 배포 파일은 [Gluon JavaFX 다운로드](https://gluonhq.com/products/javafx/)에서 확인할 수 있습니다. 다른 JavaFX 버전을 사용할 경우 Maven 의존성과 JMODs 버전을 함께 맞추세요.

```powershell
$fxJmods = 'C:\tools\javafx-jmods-21.0.4'
if (-not $env:JAVA_HOME) { throw 'JAVA_HOME을 JDK 21 설치 폴더로 설정하세요.' }
$jpackageExe = Join-Path $env:JAVA_HOME 'bin\jpackage.exe'
if (-not (Test-Path -LiteralPath $jpackageExe)) { throw 'jpackage.exe를 찾을 수 없습니다.' }
if (-not (Test-Path -LiteralPath (Join-Path $fxJmods 'javafx.controls.jmod'))) {
    throw 'JavaFX JMODs 폴더를 확인하세요.'
}

mvn clean package
if ($LASTEXITCODE -ne 0) { throw 'Maven 빌드 실패' }

# 배포에 필요한 JAR만 별도 폴더에 준비합니다.
$packageInput = '.\target\package-input'
New-Item -ItemType Directory -Path $packageInput -Force | Out-Null
Copy-Item -LiteralPath '.\target\menu-hwp-to-excel-1.0.0-all.jar' -Destination $packageInput

# 다시 실행해도 이전 배포본을 덮어쓰지 않도록 매번 새 경로를 만듭니다.
$packageStamp = [guid]::NewGuid().ToString('N')
$portableDest = Join-Path '.\portable' $packageStamp
```

### 2. 실행용 EXE 폴더 생성

```powershell
$appArgs = @(
    '--type', 'app-image',
    '--name', 'Jungwha_Converter',
    '--app-version', '1.0.0',
    '--dest', $portableDest,
    '--input', $packageInput,
    '--main-jar', 'menu-hwp-to-excel-1.0.0-all.jar',
    '--main-class', 'org.example.FxApp',
    '--icon', '.\icon.ico',
    '--module-path', "$env:JAVA_HOME\jmods;$fxJmods",
    '--add-modules', 'java.se,jdk.unsupported,jdk.charsets,javafx.controls,javafx.graphics',
    '--java-options', '--add-modules=javafx.controls,javafx.graphics',
    '--java-options', '-Dfile.encoding=UTF-8'
)
& $jpackageExe @appArgs
if ($LASTEXITCODE -ne 0) { throw '실행용 EXE 생성 실패' }

$appImage = Join-Path $portableDest 'Jungwha_Converter'
Copy-Item -LiteralPath '.\input' -Destination (Join-Path $appImage 'input') -Recurse
Write-Host "생성 위치: $appImage"
```

생성된 폴더는 다음 구조입니다.

```text
portable/<생성된 ID>/Jungwha_Converter/
├── Jungwha_Converter.exe           # 실행 파일
├── app/                           # 애플리케이션 JAR 및 설정
├── runtime/                       # Java 및 JavaFX 런타임
└── input/                         # Excel 템플릿, 이미지, 식단 예시
```

`Jungwha_Converter.exe`를 실행한 뒤 예시 HWP가 변환되는지 확인하세요. **EXE 파일만 따로 복사하면 안 됩니다.** `app`, `runtime`, `input`이 들어 있는 `Jungwha_Converter` 폴더 전체를 ZIP으로 묶어 배포합니다.

```powershell
Compress-Archive -LiteralPath $appImage -DestinationPath (Join-Path $portableDest 'Jungwha_Converter.zip')
```

### 3. 설치용 EXE 생성 (선택)

설치 마법사가 필요할 때 실행합니다. **앞 단계의 실행용 폴더 생성과 `input` 복사를 먼저 완료**해야 합니다. 이렇게 만들어 둔 앱 이미지를 설치 프로그램에 포함합니다.

JDK 21의 Windows 설치 패키징에는 WiX 도구가 필요합니다. **WiX Toolset 3.x**의 `candle.exe`, `light.exe`를 설치하고 PATH에 추가한 뒤 확인하세요. 실행용 `app-image`만 만들 때는 WiX가 필요하지 않습니다.

```powershell
Get-Command candle.exe, light.exe -ErrorAction Stop

$installerDest = Join-Path '.\dist' $packageStamp
$installerArgs = @(
    '--type', 'exe',
    '--name', 'Jungwha_Converter',
    '--app-version', '1.0.0',
    '--app-image', $appImage,
    '--dest', $installerDest,
    '--win-dir-chooser',
    '--win-menu',
    '--win-shortcut'
)
& $jpackageExe @installerArgs
if ($LASTEXITCODE -ne 0) { throw '설치용 EXE 생성 실패' }
Write-Host "설치 파일 위치: $installerDest"
```

성공하면 `dist/<생성된 ID>/`에 설치용 `.exe`가 생성됩니다. 설치한 프로그램에서 결과를 저장할 때는 문서 폴더 등 쓰기 가능한 위치를 선택하세요.

이 명령은 프로젝트 설정과 패키징 문서를 기준으로 작성한 절차입니다. 실제 배포 전에는 생성된 EXE 실행, HWP 변환, 설치 및 제거 동작을 Windows에서 확인하세요.

참고: [JDK 21 패키징 안내](https://docs.oracle.com/en/java/javase/21/jpackage/packaging-overview.html), [JavaFX 패키징 예제](https://inside.java/2023/11/14/package-javafx-native-exec/).

## 템플릿과 리소스

`input/`에는 식단 HWP 예시, 연령·유형별 조리지시서 Excel 템플릿, 이미지 리소스가 들어 있습니다.

| 구분 | 파일명 예시 |
| --- | --- |
| 일반형 기존 양식 | `2022~2025 조리지시서(만1-2세 일반형).xlsx` |
| 시간연장형 기존 양식 | `2021.9~2025 조리지시서(만1-2세 시간연장형).xlsx` |
| 2026년 일반형 양식 | `★2026~조리지시서(만1-2세 일반형).xlsx` |
| 2026년 시간연장형 양식 | `★2026~조리지시서(만1-2세 시간연장형).xlsx` |
| 이미지 | `allergy.png`, `Company.png`, `Logo.png` |

각 Excel 양식에는 만3–5세용 파일도 있습니다. 템플릿 이름은 코드에서 참조하므로 임의로 바꾸지 마세요. 기존 양식과 2026년 양식을 함께 사용하는 변환 경로가 있으므로 실행·배포 시 `input` 폴더 전체를 유지하세요.

## 프로젝트 구조

```text
javaAutomationTest/
├── input/                         # 식단 예시, 조리지시서 양식, 이미지
├── src/main/java/org/example/
│   ├── FxApp.java                 # JavaFX 진입점
│   ├── MainView.java              # 파일 선택·변환 화면
│   ├── ConverterService.java      # 백그라운드 변환 및 유형 분기
│   ├── AppHomeResolver.java       # 실행 위치와 input 경로 탐색
│   ├── HwpToJson.java             # 시간연장형 HWP 분석
│   ├── HwpToJsonGeneral.java      # 일반형 HWP 분석
│   ├── JsonToExcel.java           # 시간연장형 Excel 생성
│   ├── JsonToExcelGeneral.java    # 일반형 Excel 생성
│   ├── AllInOne.java              # 기존 실행 진입점
│   └── AllInOneUI.java            # 기존 Swing UI
├── src/main/resources/
├── icon.ico
└── pom.xml
```

GUI 변환은 `HWP 분석 → JSON 문자열 → Excel 생성` 순서로 진행됩니다. 사용자가 중간 JSON 파일을 직접 준비할 필요는 없습니다.

### 주요 라이브러리

| 역할 | 라이브러리 |
| --- | --- |
| 데스크톱 UI | JavaFX 21.0.4 |
| HWP 읽기 | hwplib 1.1.5 |
| Excel 처리 | Apache POI 5.2.5 |
| JSON 처리 | Jackson Databind 2.17.1 |
| 빌드 | Maven |

## 문제 해결

| 증상 | 확인할 내용 |
| --- | --- |
| `input 폴더가 없습니다` | 개발 실행에서는 `target/input`을 준비합니다. 패키지 실행에서는 앱의 JAR 폴더 또는 그 상위 폴더에 `input`이 있어야 합니다. |
| 템플릿을 찾지 못함 | `input` 전체가 있는지, Excel 템플릿 파일명을 바꾸지 않았는지 확인합니다. |
| 유형이나 연령이 잘못 적용됨 | HWP 파일명에 `시간연장형`, `만1-2세`, `만3-5세` 등 해당 정보를 정확히 표기합니다. |
| Excel 저장 실패 | 출력 폴더의 쓰기 권한을 확인하고, 같은 출력 파일이 Excel에서 열려 있다면 닫습니다. |
| 변환 내용이 예상과 다름 | 입력 HWP의 표 구조와 템플릿이 저장소 예시와 맞는지 확인합니다. |

생성 결과의 날짜, 메뉴, 재료 및 수량은 실제 사용 전에 확인하세요. 같은 출력 경로를 반복 사용하면 기존 결과 파일을 덮어쓸 수 있습니다.

`target/`, `portable/`, 임시 런타임 폴더와 `output/`은 Git 추적에서 제외됩니다. GUI 출력 경로는 사용자가 선택하므로 결과 파일이 반드시 `output/`에 저장되는 것은 아닙니다.
