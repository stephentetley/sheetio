use SheetReader.WbMonad;

def main(): Unit & Impure = 
    use SheetReader/WbMonad.{runWbMonad};
    match runWbMonad(test01(), System/FilePath.new("G:/work/assets/uxl/LANGS_uxl01.xlsx")) { 
        case Err(_) => Console.printLine("err")
        case Ok() => Console.printLine("[done]")
    }

def test01(): WbMonad[Unit] = 
    use SheetReader/WbMonad.{flatMap, return, getSheetNames, mapX, liftAction};
    let* xs      = getSheetNames();
    mapX(x -> liftAction(_ -> Console.printLine(x)), xs)


