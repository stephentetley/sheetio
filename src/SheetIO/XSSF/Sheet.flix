/*
 * Copyright 2026 Stephen Tetley
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


pub mod SheetIO.XSSF.Sheet {
    
    use SheetIO.Common.CellAddress.CellAddress
    use SheetIO.XSSF.Cell.Cell
    use SheetIO.XSSF.Row.Row
    
    use SheetIO.XSSF.Row.XssfRow
    use SheetIO.XSSF.Workbook.XssfWorkbook


    import org.apache.poi.xssf.usermodel.{XSSFSheet => JXSSFSheet}
    import java.lang.{Exception => JException}

    pub enum XssfSheet(JXSSFSheet)

    pub enum CellMissingPolicy with Eq, Order, ToString {
        case CREATE_NULL_AS_BLANK 
        case RETURN_BLANK_AS_NULL 
        case RETURN_NULL_AND_BLANK 
    }

    pub eff Sheet {
        def getSheetName() : String
        def getRow(i: Int32): Option[XssfRow]
        def getFirstRowNum(): Int32
        def getLastRowNum(): Int32
        def getProtect(): Bool
        def getWorkbook(): XssfWorkbook
    }

    pub def handle(sheet: XssfSheet, 
                f: a -> b \ ef): a -> Result[String, b] \ (ef - Sheet) = x -> region rc {
        handleHelper(rc, sheet, f, x)
    }


    def handleHelper(_: Region[r], 
                        sheet: XssfSheet, 
                        f: a -> b \ ef): a -> Result[String, b] \ (ef - Sheet) + r = x -> {
        let XssfSheet.XssfSheet(jsheet) = sheet;
        run {
            Ok(f(x))
        } with handler Sheet {
            
            def getSheetName(k) = try {
                let ans = unsafe IO as r { jsheet.getSheetName() };
                k(ans)
            } catch {
                case ex: JException => Err(ex.getMessage())
            }

            def getRow(i, k) = {
                let ans = unsafe IO as r { jsheet.getRow(i) };
                if (Object.isNull(ans)) 
                    k(None)
                else 
                    k(Some(XssfRow.XssfRow(ans)))
            } 

            def getFirstRowNum(k) = {
                let ans = unsafe IO as r { jsheet.getFirstRowNum() };
                k(ans)
            } 

            def getLastRowNum(k) = {
                let ans = unsafe IO as r { jsheet.getLastRowNum() };
                k(ans)
            } 
            
            def getProtect(k) = {
                let ans = unsafe IO as r { jsheet.getProtect() };
                k(ans)
            }

            def getWorkbook(k) = {
                let ans = unsafe IO as r { jsheet.getWorkbook() };
                k(XssfWorkbook.XssfWorkbook(ans))
            } 
        }
    }

    pub def runSheet(sheet: XssfSheet, f: Unit -> a \ ef): Result[String, a] \ (ef - Sheet) = 
        handle(sheet, f)()

    // TODO rewrite in a more natural style
    pub def getCellRawValue(addr: CellAddress): Option[String] \ Sheet = 
        let rix = SheetIO.Common.CellAddress.getRow(addr);
        let cix = SheetIO.Common.CellAddress.getColumn(addr);
        Option.flatMap(
            row -> ({
                    SheetIO.XSSF.Row.runRow(row, _ ->
                         Option.flatMap(
                            cell -> { 
                                SheetIO.XSSF.Cell.runCell(cell, _ -> Cell.getRawValue()) |> Result.toOption 
                                },
                            Row.getCell(cix)
                    ))
                } |> Result.toOption |> Option.flatten |> Option.flatten),
            Sheet.getRow(rix)
        )


}

