
pub mod Test01 {

    use Assert.{assertTrue}
    use SheetIO.Usermodel.ReadWorkbook.ReadWorkbook

    @Test
    pub def test01(): Unit \ { Assert, IO} = 
        let path = "./data/functional_langs.xlsx";
        discard run { 
            let n = ReadWorkbook.getNumberOfSheets();
            println("number of sheets: ${n}");
            ()
        } with SheetIO.Usermodel.ReadWorkbook.runReadWorkbookXssf(path, None);
        assertTrue(true)

}
