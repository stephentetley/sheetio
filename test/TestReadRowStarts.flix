use SheetIO.SheetReader;
use SheetIO.NamedRowEvaluator;
use SheetIO.RowEvaluator;

def main(_args: Array[String]): Int32 & Impure = 
    runtest("test01...", test01());
    runtest("test02...", test02());
    runtest("test03...", test03());
    runtest("test04...", test04());
    0

def runtest(description: String, action: SheetReader[List[(String, Int32)]]): Unit & Impure =
    use SheetIO/SheetReader.runXlsx;
    Console.printLine(description);
    match runXlsx(action, System/FilePath.new("e:/coding/flix/sheetio/data/row_starts.xlsx")) { 
        case Err(msg) => Console.printLine("Error: ${msg}")
        case Ok(xs) => List.foreach(x -> { let (a,b) = x; Console.printLine("${a}, ${b}")} , xs)
    }

def test01(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNamed, allNamedRows};      
    let* s1     = getSheetNamed("From_row_1");
    let* xs     = allNamedRows(s1, evalNamedRow());
    return(xs)

def test02(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNamed, namedRowsFrom};      
    let* s1     = getSheetNamed("From_row_3");
    let* xs     = namedRowsFrom(s1, 2, 3, evalNamedRow());
    return(xs)

def evalNamedRow(): NamedRowEvaluator[(String, Int32)] = 
    use SheetIO/NamedRowEvaluator.{tuple2, getContentNamed, getInt32Named};
    tuple2( getContentNamed("Lang")
          , getInt32Named("Year")
          )

def test03(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNamed, allRowsFrom};      
    let* s1     = getSheetNamed("From_row_1");
    let* xs     = allRowsFrom(s1, 1, evalIndexedRow());
    return(xs)

def test04(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNamed, allRowsFrom};      
    let* s1     = getSheetNamed("From_row_3");
    let* xs     = allRowsFrom(s1, 3, evalIndexedRow());
    return(xs)

def evalIndexedRow(): RowEvaluator[(String, Int32)] = 
    use SheetIO/RowEvaluator.{tuple2, getContentAt, getInt32At};
    tuple2( getContentAt(0)
          , getInt32At(1)
          )

