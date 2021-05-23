use SheetIO/SheetParser.{SheetParser1};

def main(_args: Array[String]): Int32 & Impure = 
    let src = System/FilePath.new("E:/coding/flix/sheetio/data/functional_langs.xlsx");
    match SheetIO/SheetParser.runXlsx(parser1(), "Sheet1", (), src) { 
        case Err(_) => Console.printLine("err")
        case Ok(a) => Console.printLine("OK: ${a}")
    };
    0

def parser1(): SheetParser1[(String, String)] = 
    use SheetIO/SheetParser.{flatMap, return, row, getContent};
    row({
        let* s1     = getContent();
        let* s2     = getContent();
        return((s1, s2))
    })



