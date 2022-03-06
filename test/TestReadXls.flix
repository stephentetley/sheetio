
namespace TestReadXls {

use SheetIO.SheetReader;
use SheetIO.NamedRowEvaluator;

@test
def testAll(): Int32 & Impure = 
    use SheetIO/SheetReader.{runHssf};
    match runHssf(test01(), System/FilePath.new("e:/coding/flix/sheetio/data/poi-ex-elle-ess.xls")) { 
        case Err(msg) => println("Error: ${msg}")
        case Ok(xs) => List.foreach(x -> { let (a,b) = x; println("${a}, ${b}")} , xs)
    };
    0

pub def test01(): SheetReader[List[(String, String)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allNamedRows};      
    let* s1     = getSheetNamed("Sheet1");
    let* xs     = allNamedRows(evalRow(), s1);
    point(xs)

def evalRow(): NamedRowEvaluator[(String, String)] = 
    use SheetIO/NamedRowEvaluator.{getStringNamed};
    Applicative.product(getStringNamed("Apache"), getStringNamed("Format"))

}
