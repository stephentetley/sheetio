use SheetReader.WbMonad;

def main(): Unit & Impure = 
    use SheetReader/WbMonad.{runWbMonad};
    match runWbMonad(test01(), System/FilePath.new("G:/work/assets/uxl/LANGS_uxl01.xlsx")) { 
        case Err(_) => Console.printLine("err")
        case Ok(s) => Console.printLine(s)
    }

def test01(): WbMonad[String] = 
    use SheetReader/WbMonad.{flatMap, return, getNumberOfSheets, getSheetName};
    let* i      = getNumberOfSheets();
    if (i > 0) getSheetName(0) else return("{}")
