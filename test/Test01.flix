
pub mod Test01 {

    use Assert.{assertTrue}
    use SheetIO.Workbook.Workbook

    @Test
    pub def test01(): Unit \ { Assert, IO } = 
        let path = "./data/functional_langs.xlsx";
        discard run { 
            let n = Workbook.getNumberOfSheets();
            println("number of sheets: ${n}");
            ()
        } with SheetIO.Workbook.runWithIO(path);     
        assertTrue(true)


}