
namespace TestRead01 {

    use SheetIO.WorkbookReader;
    use SheetIO.NamedRowEvaluator;
    use SheetIO.RowEvaluator;

    @test
    def testAll(): Int32 \ IO = region r {
        use SheetIO/WorkbookReader.{runXlsx};
        let path1 = Basis/NIO/File/Path.newPath("G:/work/assets/lifting-equipment-reports/_Temp01.xlsx") |> Result.getWithDefault(Basis/NIO/File/Path.empty());
        match runXlsx(r, test01(), path1) { 
            case Err(e) => println("Error: ${e}")
            case Ok(xs) => List.forEach(x -> { let (a,b) = x; println("test01: ${a}, ${b}")} , xs)
        };
        match runXlsx(r, test02(), path1) { 
            case Err(e) => println("Error: ${e}")
            case Ok(xs) => List.forEach(x -> { let (a,b) = x; println("test02: ${a}, ${b}")} , xs)
        };
        0
    }

    pub def test01(): WorkbookReader[r, ka, List[(String, String)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("Sheet1");
        let* xs     = withSheetReader(SheetIO/SheetReader.allRows(evalRow()), s1);
        point(xs)

    pub def test02(): WorkbookReader[r, ka, List[(String, String)]] \ Read(r) = 
        use SheetIO/WorkbookReader.{flatMap, point, getSheetNamed, withSheetReader};      
        let* s1     = getSheetNamed("Sheet1");
        let* xs     = withSheetReader(SheetIO/SheetReader.allNamedRowsFrom({headerIndex = 0, start = 1}, evalNamedRow()), s1);
        point(xs)


    def evalRow(): RowEvaluator[r, ka, (String, String)] \ Read(r) =
        use SheetIO/RowEvaluator.{getContentAt};
        Applicative.product(getContentAt(0), getContentAt(1))

    def evalNamedRow(): NamedRowEvaluator[r, ka, (String, String)] \ Read(r) = 
        use SheetIO/NamedRowEvaluator.{getContentNamed};
        Applicative.product(getContentNamed("Change"), getContentNamed("Reference"))          

}
