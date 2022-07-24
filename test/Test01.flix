
namespace Test01 {

    use SheetIO.WorkbookReader;

    @test
    def testAll(): Unit & Impure = 
        use SheetIO/WorkbookReader.{runXlsx};
        match runXlsx(test01(), FlixSandbox/System/FilePath.filePath("G:/work/assets/uxl/LANGS_uxl01.xlsx")) { 
            case Err(_) => println("err")
            case Ok(xs) => List.foreach(x -> println("${x.name}: ${x.start} .. ${x.end}"), xs)
        }

    def test01(): WorkbookReader[r, List[{name:: String, start :: Int32, end :: Int32}]]  \ { Read(r) } = 
        use SheetIO/WorkbookReader.{flatMap, getSheetNames};
        let* xs      = getSheetNames();
        Traversable.traverse(sheetStats, xs)

    def sheetStats(sheetName: String): WorkbookReader[r, {name:: String, start :: Int32, end :: Int32}] \ { Read(r) } = 
        use SheetIO/WorkbookReader.{flatMap, getSheetNamed, withSheetReader, point};
        use SheetIO/SheetReader.{getFirstRowNum, getLastRowNum};
        let* s1     = getSheetNamed(sheetName);
        let* i1     = withSheetReader(getFirstRowNum(), s1);
        let* i2     = withSheetReader(getLastRowNum(), s1);
        point({name = sheetName, start = i1, end = i2})


}
