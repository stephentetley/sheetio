use SheetIO.RowSetter;
use SheetIO.SheetWriter;
use SheetIO.CellStyle;

def main(_argsv: Array[String]): Int32 & Impure = 
    match SheetIO/SheetWriter.runXlsx(write1(), System/FilePath.new("E:/coding/flix/sheetio/data/bold_out1.xlsx")) { 
        case Err(e) => {Console.printLine("Error: ${e}"); 1}
        case Ok() => {Console.printLine("[writer done]"); 0}
    }
    


def write1(): SheetWriter[Unit] = 
    use SheetIO/SheetWriter.{flatMap, return, createSheet, createFont, createCellStyle, appendRow, 
            createDataFormat, getFormat};
    use SheetIO/FontSetter.{setBold!, setColor!};
    use SheetIO/FontSetter.{*>};
    use SheetIO/CellStyleSetter.{setFont!, setDataFormat!};    
    use SheetIO/IndexedColors.{getIndex}; 
    use SheetIO.IndexedColors.{Aqua}; 
    let* s1             = createSheet("Sheet1");
    let* font1          = createFont(setBold!(true) *> setColor!(getIndex(Aqua)));
    let* boldStyle      = createCellStyle(setFont!(font1));
    let* dataFmt        = createDataFormat();
    let* fmt1           = dataFmt.getFormat("0.00");
    let* numStyle       = createCellStyle(setDataFormat!(fmt1));
    let* _      = appendRow(s1, setRowOf2(boldStyle, "Hello", "World!", None));
    let* _      = appendRow(s1, setRowOf2(boldStyle, "one", "two", Some("three")));
    let* _      = appendRow(s1, setRowOf2(boldStyle, "Red", "Green", Some("Blue")));    
    let* _      = appendRow(s1, setRowOf2(boldStyle, "Old", "New", None));
    let* _      = appendRow(s1, setRowOf1(numStyle, 12.1234f64));
    let* _      = appendRow(s1, setRowOf1(numStyle, 12.0f64));
    return()
    
// Change,Reference,Asset Name,OldParentName,NewParentName
def setRowOf2(astyle: CellStyle, a: String, b: String, c: Option[String]): RowSetter[Unit] = 
    use SheetIO/RowSetter.{flatMap, return, localCellStyle, putStringAt!, putOptionStringAt!};
    let* _      = localCellStyle(astyle, putStringAt!(0, a));
    let* _      = putStringAt!(1, b);
    let* _      = putOptionStringAt!(2, c);
    return()

def setRowOf1(sty: CellStyle, a: Float64): RowSetter[Unit] = 
    use SheetIO/RowSetter.{flatMap, return, localCellStyle, putStringAt!, putFloat64At!};
    let* _      = putStringAt!(0, "Number");
    let* _      = localCellStyle(sty, putFloat64At!(1, a));
    return()
