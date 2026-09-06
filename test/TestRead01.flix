
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
        SheetIO.Usermodel.ReadWorkbook.withReadXssf(path, None, workbook -> 
            run { 
                let n = ReadWorkbook.getNumberOfSheets();
                println("number of sheets: ${n}")
            } with SheetIO.Usermodel.ReadWorkbook.runReadWorkbook(workbook)
        );
        assertTrue(true)
    
    
    @Test
    pub def readRows01(): Unit \ { Assert, IO, Fail} =
        let path = "./data/functional_langs.xlsx";
        SheetIO.Usermodel.ReadWorkbook.withReadXssf(path, None, workbook -> 
            run {
                let sheet = ReadWorkbook.getSheet("Sheet1");
                foreach(row <- sheet) {
                    run {
                        let i = ReadRow.getRowNum();
                        let name = SheetIO.Usermodel.ReadRow.readCellAt(0, ReadCell.getCellValueString);
                        let typing = SheetIO.Usermodel.ReadRow.readCellAt(1, ReadCell.getCellValueString);
                        println("${i}: ${name} ${typing}")
                    } with SheetIO.Usermodel.ReadRow.runReadRow(row)
                }
            } with SheetIO.Usermodel.ReadWorkbook.runReadWorkbook(workbook)
        );
        assertTrue(true)

        
    

}
