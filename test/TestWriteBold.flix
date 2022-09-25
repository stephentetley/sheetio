
namespace TestWriteBold {

    use SheetIO.RowSetter;
    use SheetIO.WorkbookWriter;
    use SheetIO.CellStyle;

    @test
    def testAll(): Int32 & Impure = region r {
        let path1 = Basis/System/Path.newPath("E:/coding/flix/sheetio/data/bold_out1.xlsx") |> Result.getWithDefault(Basis/System/Path.empty());
        match SheetIO/WorkbookWriter.runXlsx(r, write1(), path1) { 
            case Err(e) => {println("Error: ${e}"); 1}
            case Ok()   => {println("[writer done]"); 0}
        }
    }


    def write1(): WorkbookWriter[r, ka, Unit] \ Write(r) = 
        use Applicative.{*>, point};
        use SheetIO/WorkbookWriter.{flatMap, createSheet!, createFont!, createCellStyle!,  
                createDataFormat!, getFormat!, withSheetWriter};
        use SheetIO/FontSetter.{setBold!, setColor!};
        use SheetIO/CellStyleSetter.{setFont!, setDataFormat!};    
        use SheetIO/IndexedColors.{getIndex}; 
        use SheetIO.IndexedColors.{Aqua}; 
        let* s1             = createSheet!("Sheet1");
        let* font1          = createFont!(setBold!(true) *> setColor!(getIndex(Aqua)));
        let* boldStyle      = createCellStyle!(setFont!(font1));
        let* dataFmt        = createDataFormat!();
        let* fmt1           = getFormat!("0.00", dataFmt);
        let* numStyle       = createCellStyle!(setDataFormat!(fmt1));
        let* _      = withSheetWriter(SheetIO/SheetWriter.setRow!(0, setRowOf2(boldStyle, "Hello", "World!", None)), s1);
        let* _      = withSheetWriter(SheetIO/SheetWriter.setRow!(1, setRowOf2(boldStyle, "one", "two", Some("three"))), s1);
        let* _      = withSheetWriter(SheetIO/SheetWriter.setRow!(2, setRowOf2(boldStyle, "Red", "Green", Some("Blue"))), s1);
        let* _      = withSheetWriter(SheetIO/SheetWriter.setRow!(3, setRowOf2(boldStyle, "Old", "New", None)), s1);
        let* _      = withSheetWriter(SheetIO/SheetWriter.setRow!(4, setRowOf1(numStyle, 12.1234f64)), s1);
        let* _      = withSheetWriter(SheetIO/SheetWriter.setRow!(5, setRowOf1(numStyle, 12.0f64)), s1);
        point()
        
    // Change,Reference,Asset Name,OldParentName,NewParentName
    def setRowOf2(astyle: CellStyle[r], a: String, b: String, c: Option[String]): RowSetter[r, ka, Unit] \ Write(r) = 
        use SheetIO/RowSetter.{flatMap, point, localCellStyle, putStringAt!, putOptionStringAt!};
        let* _      = localCellStyle(astyle, putStringAt!(0, a));
        let* _      = putStringAt!(1, b);
        let* _      = putOptionStringAt!(2, c);
        point()

    def setRowOf1(sty: CellStyle[r], a: Float64): RowSetter[r, ka, Unit] \ Write(r) = 
        use SheetIO/RowSetter.{flatMap, point, localCellStyle, putStringAt!, putFloat64At!};
        let* _      = putStringAt!(0, "Number");
        let* _      = localCellStyle(sty, putFloat64At!(1, a));
        point()

}
