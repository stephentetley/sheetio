
pub mod TestRead01 {

    use Assert.{assertTrue}
    use EffectExtras.Fail
    use EffectExtras.Alt
    use SheetIO.Usermodel.ReadWorkbook
    use SheetIO.Usermodel.ReadRow
    use SheetIO.Usermodel.ReadCell

    @Test
    pub def getNumberOfSheets01(): Unit \ { Assert, IO } = 
        let path = "./data/functional_langs.xlsx";
        discard run { 
            let n = ReadWorkbook.getNumberOfSheets();
            println("number of sheets: ${n}");
            ()
        } with SheetIO.Usermodel.ReadWorkbook.runReadWorkbookXssf(path, None);
        assertTrue(true)
    
    
    @Test
    pub def readRows01(): Unit \ { Assert, IO } = region rc { 
        let path = "./data/functional_langs.xlsx";
        discard run { 
            SheetIO.Usermodel.ReadSheet.rowIterator1(rc) |> Iterator.forEach(row -> 
                run {
                    let i = ReadRow.getRowNum();
                    let name = SheetIO.Usermodel.ReadCell.runReadCellAt(0, ReadCell.getCellValueString);
                    let typesys = SheetIO.Usermodel.ReadCell.runReadCellAt(1, ReadCell.getCellValueString);
                    println("${i}: ${name} ${typesys}")
                } with SheetIO.Usermodel.ReadRow.runReadRow(row)
            )
        } with SheetIO.Usermodel.ReadSheet.runReadSheetNamed("Sheet1")
          with SheetIO.Usermodel.ReadWorkbook.runReadWorkbookXssf(path, None);
        assertTrue(true)
    }

}
