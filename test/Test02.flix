use SheetReader.WbMonad;

def main(): Unit & Impure = 
    use SheetReader/WbMonad.{runWbMonad};
    match runWbMonad(test01(), System/FilePath.new("G:/work/assets/lifting-equipment-reports/BBMED.xlsx")) { 
        case Err(_) => Console.printLine("err")
        case Ok() => Console.printLine("[done]")
    }


def test01(): WbMonad[Unit] = 
    use SheetReader/WbMonad.{flatMap, return, getSheet, getRowAt, rowEval, liftAction};
    let* s1     = getSheet("Sheet1");
    let* r1     = s1.getRowAt(0);
    let* ss     = rowEval(r1, SheetReader/RowEvaluator.getString(0));
    let* _      = liftAction(_ -> Console.printLine(ss));
    return()

