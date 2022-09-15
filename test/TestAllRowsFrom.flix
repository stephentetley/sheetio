
namespace TestAllRowsFrom {

    use SheetIO.WorkbookReader;
    use SheetIO.RowEvaluator;

    @test
    def testAll(): Int32 & Impure = 
        runtest("test 0...", test01(0));
        runtest("test 1...", test01(1));
        runtest("test 2...", test01(2));
        runtest("test 3...", test01(3));
        0

    def runtest(description: String, action: WorkbookReader[r, List[String]]): Unit & Impure =
        use SheetIO/WorkbookReader.runXlsx;
        println(description);
        let path1 = Basis/System/FilePath.filePath("e:/coding/flix/sheetio/data/one_to_ten.xlsx") |> Result.getWithDefault(Basis/System/FilePath.empty());
        match runXlsx(action, path1) { 
            case Err(msg) => println("Error: ${msg}")
            case Ok(xs) => List.foreach(println, xs)
        }

    def test01(start: Int32): WorkbookReader[r, List[String]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("one_to_ten");
        let* xs     = withSheetReader(SheetIO/SheetReader.allRowsFrom(start, evalRow()), s1);
        point(xs)

    def evalRow(): RowEvaluator[r, ka, String] \ Read(r) = SheetIO/RowEvaluator.getContentAt(0)

}
