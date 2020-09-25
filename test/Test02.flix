use SheetIO.RowEvaluator;
use SheetIO.RowSetter;
use SheetIO.SheetReader;
use SheetIO.SheetWriter;

def main(): Unit & Impure = 
    
    match SheetIO/SheetReader.runXlsx(test01(), System/FilePath.new("G:/work/assets/lifting-equipment-reports/BBMED.xlsx")) { 
        case Err(e) => Console.printLine(System/Error.toString(e))
        case Ok() => Console.printLine("[reader done]")
    };
    match SheetIO/SheetWriter.runXlsx(test02(), System/FilePath.new("E:/coding/flix/sheetio/data/out1.xlsx")) { 
        case Err(e) => Console.printLine("Error: " + System/Error.toString(e))
        case Ok() => Console.printLine("[writer done]")
    }



def test01(): SheetReader[Unit] = 
    use SheetIO/SheetReader.{flatMap, return, getSheet, getRowAt, rowEval, liftAction, 
        sheetPhysicalNumberOfRows, sheetLastRowNum, rowsInRange};
    let* _      = liftAction(_ -> Console.printLine("Running..."));        
    let* s1     = getSheet("Sheet1");
    let* r0     = s1.getRowAt(0);
    let* hdrs   = rowEval(r0, evalHeaderRow());
    let* _      = liftAction(_ -> printHeaders(hdrs));
    let* end    = sheetPhysicalNumberOfRows(s1);
    let* _      = liftAction(_ -> Console.printLine("End=${Int32.toString(end)}"));
    let* xs     = rowsInRange(s1, 1, end, evalAibExportRow());
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
    use SheetIO/RowEvaluator.{flatMap, return, getString, liftAction};
    let* s0     = getString(0);
    let* s1     = getString(1);
    let* s2     = getString(2);
    let* s3     = getString(3);
    let* s4     = getString(4);
    let* ans    = liftAction(_ -> [s0, s1, s2, s3, s4]);
    return(ans)

// Change,Reference,Asset Name,OldParentName,NewParentName
def evalAibExportRow(): RowEvaluator[AibExportRow] = 
    use SheetIO/RowEvaluator.{flatMap, return, getString, getStringWithDefault};
    let* s0     = getString(0);
    let* s1     = getString(1);
    let* s2     = getStringWithDefault(2, "");
    let* s3     = getStringWithDefault(3, "");
    let* s4     = getStringWithDefault(4, "");
    return( { change = s0
            , reference = s1
            , assetName = s2
            , oldParentName = s3
            , newParentName = s4
            } )

def printHeaders(hdrs: Headers): Unit & Impure = 
    Console.printLine(Array.formatWith(x -> x, ", ", hdrs))

def printAibExportRow(row: AibExportRow): Unit & Impure = 
    Console.printLine("{change=${row.change}, reference=${row.reference}, oldParentName=${row.oldParentName}}")


/// Writing rows with an explicit index is unpleasant...
def test02(): SheetWriter[Unit] = 
    use SheetIO/SheetWriter.{flatMap, return, createSheet, writeRow};
    let* s1     = createSheet("My_Sheet");
    let* _      = writeRow(s1, 0, setRowOf2("Hello", "World!"));
    let* _      = writeRow(s1, 1, setRowOf2("one", "two"));
    let* _      = writeRow(s1, 2, setRowOf2("Red", "Green"));
    return()
    
// Change,Reference,Asset Name,OldParentName,NewParentName
def setRowOf2(a: String, b: String): RowSetter[Unit] = 
    use SheetIO/RowSetter.{flatMap, return, putString!};
    let* _      = putString!(0, a);
    let* _      = putString!(1, b);
    return()
