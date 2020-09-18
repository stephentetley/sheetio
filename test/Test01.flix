

def main(): Unit & Impure = 
    use SheetReader/Internal/Prim.{newXSSFWorkbook, getNumberOfSheets, closeXSSFWorkbook};
    use Result.{flatMap};
    let step = () -> { 
        let* wb     = newXSSFWorkbook("G:/work/assets/uxl/LANGS_uxl01.xlsx");
        let* i      = getNumberOfSheets(wb);
        let* _      = closeXSSFWorkbook(wb);
        Ok(i)
    };
    match step() { 
        case Err(_) => Console.printLine("err")
        case Ok(i) => Console.printLine("Sheet count: " + Int32.toString(i))
    }