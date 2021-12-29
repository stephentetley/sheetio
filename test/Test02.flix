use SheetIO.RowEvaluator;
use SheetIO.RowSetter;
use SheetIO.SheetReader;
use SheetIO.SheetWriter;

def main(): Unit & Impure = 
    
    match SheetIO/SheetReader.runXlsx(test01(), System/FilePath.new("G:/work/assets/lifting-equipment-reports/BBMED.xlsx")) { 
        case Err(e) => println(System/Error.toString(e))
        case Ok() => println("[reader done]")
    };
    match SheetIO/SheetWriter.runXlsx(test02(), System/FilePath.new("E:/coding/flix/sheetio/data/out1.xlsx")) { 
        case Err(e) => println("Error: " + System/Error.toString(e))
        case Ok() => println("[writer done]")
    }



def test01(): SheetReader[Unit] = 
    use SheetIO/SheetReader.{flatMap, return, getSheet, getRowAt, rowEval, liftAction, 
        sheetPhysicalNumberOfRows, sheetLastRowNum, rowsInRange};
    let* _      = liftAction(_ -> println("Running..."));        
    let* s1     = getSheet("Sheet1");
    let* r0     = s1.getRowAt(0);
    let* hdrs   = rowEval(r0, evalHeaderRow());
    let* _      = liftAction(_ -> printHeaders(hdrs));
    let* end    = sheetPhysicalNumberOfRows(s1);
    let* _      = liftAction(_ -> println("End=${Int32.toString(end)}"));
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
def evalHeaderRow(): RowEvaluator[Headers] = SheetIO/RowEvaluator.getHeaders()

// Change,Reference,Asset Name,OldParentName,NewParentName
def evalAibExportRow(): RowEvaluator[AibExportRow] = 
    use SheetIO/RowEvaluator.{flatMap, return, getStringAt, getStringAtWithDefault};
    let* s0     = getStringAt(0);
    let* s1     = getStringAt(1);
    let* s2     = getStringAtWithDefault(2, "");
    let* s3     = getStringAtWithDefault(3, "");
    let* s4     = getStringAtWithDefault(4, "");
    return( { change = s0
            , reference = s1
            , assetName = s2
            , oldParentName = s3
            , newParentName = s4
            } )

def printHeaders(hdrs: Headers): Unit & Impure = 
    println(Array.formatWith(x -> x, ", ", hdrs))

def printAibExportRow(row: AibExportRow): Unit & Impure = 
    println("{change=${row.change}, reference=${row.reference}, oldParentName=${row.oldParentName}}")



def test02(): SheetWriter[Unit] = 
    use SheetIO/SheetWriter.{flatMap, return, createSheet, appendRow};
    let* s1     = createSheet("My_Sheet");
    let* _      = appendRow(s1, setRowOf2("Hello", "World!"));
    let* _      = appendRow(s1, setRowOf2("one", "two"));
    let* _      = appendRow(s1, setRowOf2("Red", "Green"));    
    let* _      = appendRow(s1, setRowOf2("Old", "New"));
    return()
    
// Change,Reference,Asset Name,OldParentName,NewParentName
def setRowOf2(a: String, b: String): RowSetter[Unit] = 
    use SheetIO/RowSetter.{flatMap, return, putStringAt!};
    let* _      = putStringAt!(0, a);
    let* _      = putStringAt!(1, b);
    return()
