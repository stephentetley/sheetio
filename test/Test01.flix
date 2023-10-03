
mod Test01 {

    use Monad.{>>=}
    use IOError.{Generic}
    use SheetIO.WorkbookReader

    def runResult(fn: Unit -> Result[IOError, Bool] \ IO): Bool \ IO = 
        match fn() {
            case Ok(b)  => b
            case Err(e) => {println("${e}"); false}
        }

    @test
    def test01(): Bool \ IO = region rc {
        use SheetIO.WorkbookReader.{runXlsx};
        runResult(_ -> 
            Result.mapErr(Generic, Basis.IO.Path.new("E:/coding/flix/sheetio/data/functional_langs.xlsx")) >>= path ->
            runXlsx(rc, reader01(), path) >>= vec -> 
            {Vector.forEach(println, vec); Ok(true)}
        )
    }

    def reader01(): WorkbookReader[ka, r, ef + r, Vector[String]] = 
        SheetIO.WorkbookReader.readSheetNamed("Sheet1", SheetIO.SheetReader.getHeaders(0))

    @test
    def test02(): Bool \ IO = region rc {
        use SheetIO.WorkbookReader.{runXlsx};
        runResult(_ -> 
            Result.mapErr(Generic, Basis.IO.Path.new("E:/coding/flix/sheetio/data/functional_langs.xlsx")) >>= path -> 
            runXlsx(rc, reader02(), path) >>= rows ->  
            {List.forEach(println, rows); Ok(true)}
        )
    }


    def reader02(): WorkbookReader[ka, r, ef + r, List[(String, String)]] = 
        use SheetIO.RowReader.{readCellAt};
        use SheetIO.CellReader.{getTrimmedContent};
        let eval1 = GradedApplicative.product(readCellAt(0, getTrimmedContent()), readCellAt(1, getTrimmedContent()));
        SheetIO.WorkbookReader.readSheetNamed("Sheet1", SheetIO.SheetReader.allRows(eval1))


}
