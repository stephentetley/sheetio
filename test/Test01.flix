
mod Test01 {

    use GradedMonad.{>>=}
    use SheetIO.WorkbookReader

    @test
    def test01(): Bool \ IO = region rc {
        use SheetIO.WorkbookReader.{runXlsx};
        let path1 = Basis.NIO.File.Path.newPath("E:/coding/flix/sheetio/data/functional_langs.xlsx") |> Result.getWithDefault(Basis.NIO.File.Path.empty());
        match runXlsx(rc, reader01(), path1) { 
            case Err(msg) => {println("err: ${msg}"); false}
            case Ok(v)    => {Vector.forEach(println, v); true}
        }
    }

    def reader01(): WorkbookReader[ka, r, ef + r, Vector[String]] = 
        SheetIO.WorkbookReader.readSheetNamed("Sheet1", SheetIO.SheetReader.getHeaders(0))


}
