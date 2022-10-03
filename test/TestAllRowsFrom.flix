
namespace TestAllRowsFrom {

    use SheetIO.WorkbookReader;
    use SheetIO.RowEvaluator;

    @test
    def testAll(): Int32 & Impure = region r {
        runtest(r, "test 0...", test01(0));
        runtest(r, "test 1...", test01(1));
        runtest(r, "test 2...", test01(2));
        runtest(r, "test 3...", test01(3));
        0
    }

    def runtest(r: Region[r],  description: String, action: WorkbookReader[r, List[String], List[String]]): Unit \ IO =
        use SheetIO/WorkbookReader.runXlsx;
        println(description);
        let path1 = Basis/NIO/File/Path.newPath("e:/coding/flix/sheetio/data/one_to_ten.xlsx") |> Result.getWithDefault(Basis/NIO/File/Path.empty());
        match runXlsx(r, action, path1) { 
            case Err(msg) => println("Error: ${msg}")
            case Ok(xs) => List.foreach(println, xs)
        }

    def test01(start: Int32): WorkbookReader[r, ka, List[String]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("one_to_ten");
        let* xs     = withSheetReader(SheetIO/SheetReader.allRowsFrom(start, evalRow()), s1);
        point(xs)

    def evalRow(): RowEvaluator[r, ka, String] \ Read(r) = SheetIO/RowEvaluator.getContentAt(0)

}
