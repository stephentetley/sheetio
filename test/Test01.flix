
pub mod Test01 {

    use Assert.{assertTrue}
    use SheetIO.XSSF.Workbook.Workbook

    @Test
    pub def test01(): Unit \ { Assert, IO } = 
        let path = "./data/functional_langs.xlsx";
        discard run { 
            let n = Workbook.getNumberOfSheets();
            println("number of sheets: ${n}");
            ()
        } with SheetIO.XSSF.Workbook.runWithIO(path);     
        assertTrue(true)


}