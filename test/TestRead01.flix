
namespace TestRead01 {

use SheetIO.SheetReader;
use SheetIO.NamedRowEvaluator;
use SheetIO.RowEvaluator;

@test
def testAll(): Int32 & Impure = 
    use SheetIO/SheetReader.{runXlsx};
    match runXlsx(test01(), System/FilePath.new("G:/work/assets/lifting-equipment-reports/_Temp01.xlsx")) { 
        case Err(e) => println("Error: ${e}")
        case Ok(xs) => List.foreach(x -> { let (a,b) = x; println("test01: ${a}, ${b}")} , xs)
    };
    match runXlsx(test02(), System/FilePath.new("G:/work/assets/lifting-equipment-reports/_Temp01.xlsx")) { 
        case Err(e) => println("Error: ${e}")
        case Ok(xs) => List.foreach(x -> { let (a,b) = x; println("test02: ${a}, ${b}")} , xs)
    };
    0

pub def test01(): SheetReader[List[(String, String)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allRows};      
    let* s1     = getSheetNamed("Sheet1");
    let* xs     = allRows(evalRow(), s1);
    point(xs)

pub def test02(): SheetReader[List[(String, String)]] = 
    use SheetIO/SheetReader.{flatMap, point, getSheetNamed, namedRowsFrom};      
    let* s1     = getSheetNamed("Sheet1");
    let* xs     = namedRowsFrom({headerIndex = 0, start = 1}, evalNamedRow(), s1);
    point(xs)


def evalRow(): RowEvaluator[(String, String)] =
    use SheetIO/RowEvaluator.{getContentAt};
    Applicative.product(getContentAt(0), getContentAt(1))

def evalNamedRow(): NamedRowEvaluator[(String, String)] = 
    use SheetIO/NamedRowEvaluator.{getContentNamed};
    Applicative.product(getContentNamed("Change"), getContentNamed("Reference"))          

}
