
namespace TestReadRowStarts {

    use SheetIO.WorkbookReader;
    use SheetIO.NamedRowEvaluator;
    use SheetIO.RowEvaluator;

    @test
    def testAll(): Int32 \ IO = region r {
        runtest(r, "test01...", test01());
        runtest(r, "test02...", test02());
        runtest(r, "test03...", test03());
        runtest(r, "test04...", test04());
        0
    }

    def runtest(r: Region[r], description: String, action: WorkbookReader[r, List[(String, Int32)], List[(String, Int32)]]): Unit & Impure =
        use SheetIO/WorkbookReader.runXlsx;
        println(description);
        let path1 = Basis/NIO/File/Path.newPath("e:/coding/flix/sheetio/data/row_starts.xlsx") |> Result.getWithDefault(Basis/NIO/File/Path.empty());
        match runXlsx(r, action, path1) { 
            case Err(msg) => println("Error: ${msg}")
            case Ok(xs) => List.foreach(x -> { let (a,b) = x; println("${a}, ${b}")} , xs)
        }

    def test01(): WorkbookReader[r, ka, List[(String, Int32)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("From_row_1");
        let* xs     = withSheetReader(SheetIO/SheetReader.allNamedRows(evalNamedRow()), s1);
        point(xs)

    def test02(): WorkbookReader[r, ka, List[(String, Int32)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("From_row_3");
        let* xs     = withSheetReader(SheetIO/SheetReader.allNamedRowsFrom({headerIndex = 2, start = 3}, evalNamedRow()), s1);
        point(xs)

    def evalNamedRow(): NamedRowEvaluator[r, ka, (String, Int32)] \ Read(r) = 
        use SheetIO/NamedRowEvaluator.{getContentNamed, getInt32Named};
        Applicative.product(getContentNamed("Lang"), getInt32Named("Year"))

    def test03(): WorkbookReader[r, ka, List[(String, Int32)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("From_row_1");
        let* xs     = withSheetReader(SheetIO/SheetReader.allRowsFrom(1, evalIndexedRow()), s1);
        point(xs)

    def test04(): WorkbookReader[r, ka, List[(String, Int32)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("From_row_3");
        let* xs     = withSheetReader(SheetIO/SheetReader.allRowsFrom(3, evalIndexedRow()), s1);
        point(xs)

    def evalIndexedRow(): RowEvaluator[r, ka, (String, Int32)] \ Read(r) = 
        use SheetIO/RowEvaluator.{getContentAt, getInt32At};
        Applicative.product(getContentAt(0), getInt32At(1))

}