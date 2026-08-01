
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
        discard SheetIO.Usermodel.ReadWorkbook.runReadWorkbookXssf(path, None, _ -> { 
            let sheet = ReadWorkbook.getSheet("Sheet1");
            Iterable.iterator(rc, sheet) |> Iterator.forEach(row -> {
                SheetIO.Usermodel.ReadRow.runReadRow(row, _ -> {
                    let i = ReadRow.getRowNum();
                    let c1 = ReadRow.getCell(0);
                    let c2 = ReadRow.getCell(1);
                    let name = SheetIO.Usermodel.ReadCell.runReadCell(c1, ReadCell.getCellValueString);
                    let typesys = SheetIO.Usermodel.ReadCell.runReadCell(c2, ReadCell.getCellValueString);
                    println("${i}: ${name} ${typesys}")
                })
            })
        });
        assertTrue(true)
    }

}
