use SheetIO.SheetReader;
use SheetIO.RowEvaluator;

def main(_args: Array[String]): Int32 & Impure = 
    runtest("test 0...", test01(0));
    runtest("test 1...", test01(1));
    runtest("test 2...", test01(2));
    runtest("test 3...", test01(3));
    0

def runtest(description: String, action: SheetReader[List[String]]): Unit & Impure =
    use SheetIO/SheetReader.runXlsx;
    Console.printLine(description);
    match runXlsx(action, System/FilePath.new("e:/coding/flix/sheetio/data/one_to_ten.xlsx")) { 
        case Err(msg) => Console.printLine("Error: ${msg}")
        case Ok(xs) => List.foreach(println, xs)
    }

def test01(start: Int32): SheetReader[List[String]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheetNamed, allRowsFrom};      
    let* s1     = getSheetNamed("one_to_ten");
    let* xs     = allRowsFrom(s1, start, evalRow());
    return(xs)

def evalRow(): RowEvaluator[String] = 
    SheetIO/RowEvaluator.getContentAt(0)

