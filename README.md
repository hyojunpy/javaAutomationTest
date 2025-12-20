jpackage --type exe --name Jungwha_Converter --app-version 1.0.0 --input target --main-jar menu-hwp-to-excel-1.0.0-all.jar --main-class org.example.AllInOne --dest dist --java-options "-Dfile.encoding=UTF-8"

jpackage --type app-image --name Jungwha_Converter --app-version 1.0.0 --input target --main-jar menu-hwp-to-excel-1.0.0-all.jar --main-class org.example.AllInOne --dest dist --java-options "-Dfile.encoding=UTF-8"

mvn -U clean package; jpackage --type app-image --name Jungwha_Converter --dest portable --input target --main-jar menu-hwp-to-excel-1.0.0-all.jar --main-class org.example.FxApp --icon goorem2.ico; if (!(Test-Path portable\Jungwha_Converter\input)) { New-Item -ItemType Directory -Path portable\Jungwha_Converter\input | Out-Null }; Copy-Item -Recurse -Force .\input\* portable\Jungwha_Converter\input\
