
namespace Test01 {

    use SheetIO.WorkbookReader;

    @test
    def testAll(): Unit \ IO = region r {
        use SheetIO/WorkbookReader.{runXlsx};
        let path1 = Basis/NIO/File/Path.newPath("G:/work/assets/uxl/LANGS_uxl01.xlsx") |> Result.getWithDefault(Basis/NIO/File/Path.empty());
        match runXlsx(r, test01(), path1) { 
            case Err(_) => println("err")
            case Ok(xs) => List.forEach(x -> println("${x.name}: ${x.start} .. ${x.end}"), xs)
        }
    }

    def test01(): WorkbookReader[r, ka, List[{name = String, start = Int32, end = Int32}]]  \ { Read(r) } = 
        use SheetIO/WorkbookReader.{flatMap, getSheetNames};
        let* xs      = getSheetNames();
        Traversable.traverse(sheetStats, xs)

    def sheetStats(sheetName: String): WorkbookReader[r, ka, {name = String, start = Int32, end = Int32}] \ { Read(r) } = 
        use SheetIO/WorkbookReader.{flatMap, getSheetNamed, withSheetReader, point};
        use SheetIO/SheetReader.{getFirstRowNum, getLastRowNum};
        let* s1     = getSheetNamed(sheetName);
        let* i1     = withSheetReader(getFirstRowNum(), s1);
        let* i2     = withSheetReader(getLastRowNum(), s1);
        point({name = sheetName, start = i1, end = i2})


}
