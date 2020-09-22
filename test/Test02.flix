use SheetReader.RowEvaluator;
use SheetReader.WbMonad;

def main(): Unit & Impure = 
    use SheetReader/WbMonad.{runWbMonad};
    match runWbMonad(test01(), System/FilePath.new("G:/work/assets/lifting-equipment-reports/BBMED.xlsx")) { 
        case Err(_) => Console.printLine("err")
        case Ok() => Console.printLine("[done]")
    }


def test01(): WbMonad[Unit] = 
    use SheetReader/WbMonad.{flatMap, return, getSheet, getRowAt, rowEval, liftAction, sheetLastRowNum, rowsInRange};
    let* s1     = getSheet("Sheet1");
    let* r0     = s1.getRowAt(0);
    let* hdrs   = rowEval(r0, evalHeaderRow());
    let* _      = liftAction(_ -> printHeaders(hdrs));
    let* end    = sheetLastRowNum(s1);
    let* _      = liftAction(_ -> Console.printLine("End=${Int32.toString(end)}"));
    let* xs     = rowsInRange(s1, 1, 4, evalAibExportRow());
    let* _      = liftAction(_ -> List.foreach(printAibExportRow, xs));
    return()

type alias Headers = Array[String]

type alias AibExportRow = 
    { change: String
    , reference: String
    , assetName: String
    , oldParentName: String
    , newParentName: String
    }

// Change,Reference,Asset Name,OldParentName,NewParentName
def evalHeaderRow(): RowEvaluator[Headers] = 
    use SheetReader/RowEvaluator.{flatMap, return, getString, liftAction};
    let* s0     = getString(0);
    let* s1     = getString(1);
    let* s2     = getString(2);
    let* s3     = getString(3);
    let* s4     = getString(4);
    let* ans    = liftAction(_ -> [s0, s1, s2, s3, s4]);
    return(ans)

// Change,Reference,Asset Name,OldParentName,NewParentName
def evalAibExportRow(): RowEvaluator[AibExportRow] = 
    use SheetReader/RowEvaluator.{flatMap, return, getString, liftAction};
    let* s0     = getString(0);
    let* s1     = getString(1);
    let* s2     = getString(2);
    let* s3     = getString(3);
    let* s4     = getString(4);
    return( { change = s0
            , reference = s1
            , assetName = s2
            , oldParentName = s3
            , newParentName = s4
            } )

def printHeaders(hdrs: Headers): Unit & Impure = 
    Console.printLine(Array.formatWith(x -> x, ", ", hdrs))

def printAibExportRow(row: AibExportRow): Unit & Impure = 
    Console.printLine("{change=${row.change}, reference=${row.reference}}")