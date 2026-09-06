
pub mod TestRead01 {

    use Assert.{assertTrue}
    use EffectExtras.Fail
    use EffectExtras.Alt
    use SheetIO.Usermodel.ReadWorkbook
    use SheetIO.Usermodel.ReadRow
    use SheetIO.Usermodel.ReadCell

    @Test
    pub def getNumberOfSheets01(): Unit \ { Assert, IO, Fail } = 
        let path = "./data/functional_langs.xlsx";
        SheetIO.Usermodel.ReadWorkbook.withWorkbookXssf(path, None, workbook -> 
            run { 
                let n = ReadWorkbook.getNumberOfSheets();
                println("number of sheets: ${n}")
            } with SheetIO.Usermodel.ReadWorkbook.runReadWorkbookXssf(workbook)
        );
        assertTrue(true)
    
    
    @Test
    pub def readRows01(): Unit \ { Assert, IO, Fail} =
        let path = "./data/functional_langs.xlsx";
        SheetIO.Usermodel.ReadWorkbook.withWorkbookXssf(path, None, workbook -> 
            foreach(sheet <- workbook) {
                foreach(row <- sheet) {
                    run {
                        let i = ReadRow.getRowNum();
                        let name = SheetIO.Usermodel.ReadCell.runReadCellAt(0, ReadCell.getCellValueString);
                        let typesys = SheetIO.Usermodel.ReadCell.runReadCellAt(1, ReadCell.getCellValueString);
                        println("${i}: ${name} ${typesys}")
                    } with SheetIO.Usermodel.ReadRow.runReadRow(row)
                }
            }
        );
        assertTrue(true)

        
    

}
