use SheetIO.SheetReader;

def main(): Unit & Impure = 
    use SheetIO/SheetReader.{runXlsx};
    match runXlsx(test01(), System/FilePath.new("G:/work/assets/uxl/LANGS_uxl01.xlsx")) { 
        case Err(_) => Console.printLine("err")
        case Ok() => Console.printLine("[done]")
    }

def sheeetStats(sheetName: String): SheetReader[Unit] = 
    use SheetIO/SheetReader.{flatMap, return, getSheet, sheetFirstRowNum, sheetLastRowNum, liftAction};
    let* s1     = getSheet(sheetName);
    let* i1     = sheetFirstRowNum(s1);
    let* i2     = sheetLastRowNum(s1);
    liftAction(_ -> Console.printLine("${sheetName}: rows=${Int32.toString(i1)}..${Int32.toString(i2)}"))


def test01(): SheetReader[Unit] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNames, mapX, liftAction};
    let* xs      = getSheetNames();
    mapX(sheeetStats, xs)


