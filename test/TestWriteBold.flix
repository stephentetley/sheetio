use SheetIO.RowSetter;
use SheetIO.SheetWriter;
use SheetIO.CellStyle;

def main(): Unit & Impure = 
    match SheetIO/SheetWriter.runXlsx(write1(), System/FilePath.new("E:/coding/flix/sheetio/data/bold_out1.xlsx")) { 
        case Err(e) => Console.printLine("Error: " + System/Error.toString(e))
        case Ok() => Console.printLine("[writer done]")
    }


def write1(): SheetWriter[Unit] = 
    use SheetIO/SheetWriter.{flatMap, return, createSheet, createFont, createCellStyle, appendRow};
    use SheetIO/FontSetter.{setBold!};
    use SheetIO/CellStyleSetter.{setFont!};
    let* s1             = createSheet("Sheet1");
    let* font1          = createFont(setBold!(true));
    let* boldStyle      = createCellStyle(setFont!(font1));
    let* _      = appendRow(s1, setRowOf2(boldStyle, "Hello", "World!"));
    let* _      = appendRow(s1, setRowOf2(boldStyle, "one", "two"));
    let* _      = appendRow(s1, setRowOf2(boldStyle, "Red", "Green"));    
    let* _      = appendRow(s1, setRowOf2(boldStyle, "Old", "New"));
    return()
    
// Change,Reference,Asset Name,OldParentName,NewParentName
def setRowOf2(astyle: CellStyle, a: String, b: String): RowSetter[Unit] = 
    use SheetIO/RowSetter.{flatMap, return, putCellStyleAt!, putStringAt!};
    let* _      = putCellStyleAt!(0, astyle);
    let* _      = putStringAt!(0, a);
    let* _      = putStringAt!(1, b);
    return()
