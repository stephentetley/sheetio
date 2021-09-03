use SheetIO.SheetReader;
use SheetIO.NamedRowEvaluator;

def main(_args: Array[String]): Int32 & Impure = 
    runtest("test01...", test01());
    runtest("test02...", test02());
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
    let* xs     = allNamedRows(s1, evalRow());
    return(xs)

def test02(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNamed, namedRowsFrom};      
    let* s1     = getSheetNamed("From_row_3");
    let* xs     = namedRowsFrom(s1, 2, 3, evalRow());
    return(xs)

def evalRow(): NamedRowEvaluator[(String, Int32)] = 
    use SheetIO/NamedRowEvaluator.{tuple2, getContentNamed, getInt32Named};
    tuple2( getContentNamed("Lang")
          , getInt32Named("Year")
          )


