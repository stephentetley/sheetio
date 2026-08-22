

pub mod TestWrite01 {
    use SheetIO.Usermodel.ReadWorkbook.sheetIterator

    use Assert.{assertEq}
    use SheetIO.Usermodel.WriteWorkbook
    use SheetIO.Usermodel.WriteSheet

    @Test
    pub def test01(): Unit \ { Assert, IO } = 
        let path = "./output/new_sheet.xlsx";
        let ans = run { 
            let sheet1 = WriteWorkbook.createSheetWithName("sheet_no1");
            SheetIO.Usermodel.WriteSheet.runWriteSheet(sheet1, _ -> {
                let _row = WriteSheet.createRow(1);
                10
            })
        } with SheetIO.Usermodel.WriteWorkbook.runWriteXssf(path);
        assertEq(expected = Ok(10), ans)

}