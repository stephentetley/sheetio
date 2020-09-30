use SheetIO.SheetReader;
use SheetIO.NamedRowEvaluator;

def main(): Unit & Impure = 
    use SheetIO/SheetReader.{runHssf};
    match runHssf(test01(), System/FilePath.new("e:/coding/flix/sheetio/data/poi-ex-elle-ess.xls")) { 
        case Err(_) => Console.printLine("err")
        case Ok(xs) => List.foreach(x -> { let (a,b) = x; Console.printLine("${a}, ${b}")} , xs)
    }

pub def test01(): SheetReader[List[(String, String)]] = 
    use SheetIO/SheetReader.{flatMap, return, getSheet, allNamedRows};      
    let* s1     = getSheet("Sheet1");
    let* xs     = allNamedRows(s1, evalRow());
    return(xs)

def evalRow(): NamedRowEvaluator[(String, String)] = 
    use SheetIO/NamedRowEvaluator.{tuple2, getStringNamed};
    tuple2( getStringNamed("Apache")
          , getStringNamed("Format")
          )


