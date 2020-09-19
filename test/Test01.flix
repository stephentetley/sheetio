use SheetReader.WbMonad;

def main(): Unit & Impure = 
    use SheetReader/WbMonad.{runWbMonad};
    match runWbMonad(test01(), System/FilePath.new("G:/work/assets/uxl/LANGS_uxl01.xlsx")) { 
        case Err(_) => Console.printLine("err")
        case Ok() => Console.printLine("[done]")
    }

def sheeetStats(sheetName: String): WbMonad[Unit] = 
    use SheetReader/WbMonad.{flatMap, return, getSheet, sheetLastRowNum, liftAction};
    let* s1     = getSheet(sheetName);
    let* i      = sheetLastRowNum(s1);
    liftAction(_ -> Console.printLine("${sheetName}: rows=" + Int32.toString(i)))


def test01(): WbMonad[Unit] = 
    use SheetReader/WbMonad.{flatMap, return, getSheetNames, mapX, liftAction};
    let* xs      = getSheetNames();
    mapX(sheeetStats, xs)


