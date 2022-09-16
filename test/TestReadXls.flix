
namespace TestReadXls {


    use SheetIO.WorkbookReader;
    use SheetIO.NamedRowEvaluator;

    @test
    def testAll(): Int32 & Impure = 
        use SheetIO/WorkbookReader.{runHssf};
        let path1 = Basis/System/FilePath.filePath("e:/coding/flix/sheetio/data/poi-ex-elle-ess.xls") |> Result.getWithDefault(Basis/System/FilePath.empty());
        match runHssf(test01(), path1) { 
            case Err(msg) => println("Error: ${msg}")
            case Ok(xs) => List.foreach(x -> { let (a,b) = x; println("${a}, ${b}")} , xs)
        };
        0

    pub def test01(): WorkbookReader[r, List[(String, String)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};
        let* s1     = getSheetNamed("Sheet1");
        let* xs     = withSheetReader(SheetIO/SheetReader.allNamedRows(evalRow()), s1);
        point(xs)

    def evalRow(): NamedRowEvaluator[r, ka, (String, String)] \ Read(r) = 
        use SheetIO/NamedRowEvaluator.{getStringNamed};
        Applicative.product(getStringNamed("Apache"), getStringNamed("Format"))

}
