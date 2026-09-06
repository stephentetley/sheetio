

pub mod TestWrite01 {
    
    use Assert.{assertEq}
    use EffectExtras.Fail
    use SheetIO.Usermodel.WriteWorkbook
    use SheetIO.Usermodel.WriteSheet

    @Test
    pub def test01(): Unit \ { Assert, IO, Fail } = 
        let path = "./output/new_sheet.xlsx";
        let ans = SheetIO.Usermodel.WriteWorkbook.withWriteXssf(path, workbook -> {
            run { 
                let sheet1 = WriteWorkbook.createSheetWithName("sheet_no1");
                SheetIO.Usermodel.WriteSheet.runWriteSheet(sheet1, _ -> {
                    let _row = WriteSheet.createRow(1);
                    10
                })
            } with SheetIO.Usermodel.WriteWorkbook.runWriteWorkbook(workbook)
        });
        assertEq(expected = 10, ans)

}