
namespace Test01 {

use SheetIO.SheetReader;

@test
def testAll(): Unit & Impure = 
    use SheetIO/SheetReader.{runXlsx};
    match runXlsx(test01(), FlixSandbox/System/FilePath.new("G:/work/assets/uxl/LANGS_uxl01.xlsx")) { 
        case Err(_) => println("err")
        case Ok() => println("[done]")
    }

def sheeetStats(sheetName: String): SheetReader[Unit] = 
    use SheetIO/SheetReader.{flatMap, getSheetNamed, getFirstRowNum, getLastRowNum, liftAction};
    let* s1     = getSheetNamed(sheetName);
    let* i1     = getFirstRowNum(s1);
    let* i2     = getLastRowNum(s1);
    liftAction(_ -> println("${sheetName}: rows=${Int32.toString(i1)}..${Int32.toString(i2)}"))


def test01(): SheetReader[Unit] = 
    use SheetIO/SheetReader.{flatMap, getSheetNames, mapX};
    let* xs      = getSheetNames();
    mapX(sheeetStats, xs)

}
