
namespace TestAllRowsFrom {

use SheetIO.SheetReader;
use SheetIO.RowEvaluator;

@test
def testAll(): Int32 & Impure = 
    runtest("test 0...", test01(0));
    runtest("test 1...", test01(1));
    runtest("test 2...", test01(2));
    runtest("test 3...", test01(3));
    0

def runtest(description: String, action: SheetReader[List[String]]): Unit & Impure =
    use SheetIO/SheetReader.runXlsx;
    println(description);
    match runXlsx(action, FlixSandbox/System/FilePath.new("e:/coding/flix/sheetio/data/one_to_ten.xlsx")) { 
        case Err(msg) => println("Error: ${msg}")
        case Ok(xs) => List.foreach(println, xs)
    }

def test01(start: Int32): SheetReader[List[String]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allRowsFrom};      
    let* s1     = getSheetNamed("one_to_ten");
    let* xs     = allRowsFrom(start, evalRow(), s1);
    point(xs)

def evalRow(): RowEvaluator[String] = 
    SheetIO/RowEvaluator.getContentAt(0)


}