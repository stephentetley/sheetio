
set %JAVA_HOME%="C:\Program Files\Zulu\zulu-11"

javac -cp "../lib/poi-5.0.0.jar;../lib/poi-ooxml-5.0.0.jar;../lib/poi-ooxml-full-5.0.0.jar;../lib/excel-streaming-reader-3.1.1.jar" -d ./build  flixspt/sheetio/*.java

cd build 

jar cf sheetiojava-1.0.jar ./flixspt/sheetio/*.class

