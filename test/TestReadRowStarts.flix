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
    println(description);
    match runXlsx(action, System/FilePath.new("e:/coding/flix/sheetio/data/row_starts.xlsx")) { 
        case Err(msg) => println("Error: ${msg}")
        case Ok(xs) => List.foreach(x -> { let (a,b) = x; println("${a}, ${b}")} , xs)
    }

def test01(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allNamedRows};      
    let* s1     = getSheetNamed("From_row_1");
    let* xs     = allNamedRows(evalNamedRow(), s1);
    point(xs)

def test02(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, namedRowsFrom};      
    let* s1     = getSheetNamed("From_row_3");
    let* xs     = namedRowsFrom(2, 3, evalNamedRow(), s1);
    point(xs)

def evalNamedRow(): NamedRowEvaluator[(String, Int32)] = 
    use SheetIO/NamedRowEvaluator.{getContentNamed, getInt32Named};
    Applicative.product(getContentNamed("Lang"), getInt32Named("Year"))

def test03(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allRowsFrom};      
    let* s1     = getSheetNamed("From_row_1");
    let* xs     = allRowsFrom(1, evalIndexedRow(), s1);
    point(xs)

def test04(): SheetReader[List[(String, Int32)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allRowsFrom};      
    let* s1     = getSheetNamed("From_row_3");
    let* xs     = allRowsFrom(3, evalIndexedRow(), s1);
    point(xs)

def evalIndexedRow(): RowEvaluator[(String, Int32)] = 
    use SheetIO/RowEvaluator.{getContentAt, getInt32At};
    Applicative.product(getContentAt(0), getInt32At(1))

