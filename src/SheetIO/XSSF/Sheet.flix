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
    use SheetIO.XSSF.Cell.XssfCell
    use SheetIO.XSSF.Row.Row    
    use SheetIO.XSSF.Row.XssfRow
    use SheetIO.XSSF.Workbook.XssfWorkbook


    import org.apache.poi.xssf.usermodel.{XSSFSheet => JXSSFSheet}
    import org.apache.poi.xssf.usermodel.{XSSFRow => JXSSFRow}
    import org.apache.poi.xssf.usermodel.{XSSFCell => JXSSFCell}
    import org.apache.poi.ss.usermodel.{Row => JRow} // superclass
    import org.apache.poi.ss.util.{SheetUtil => JSheetUtil}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}

    pub enum XssfSheet(JXSSFSheet)

    pub enum CellMissingPolicy with Eq, Order, ToString {
        case CREATE_NULL_AS_BLANK 
        case RETURN_BLANK_AS_NULL 
        case RETURN_NULL_AND_BLANK 
    }

    pub eff Sheet {
        // `askSheetHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askSheetHandle(): XssfSheet
        def getSheetName() : String
        def getRow(i: Int32): Option[XssfRow]
        def getCell(addr: CellAddress): Option[XssfCell]
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
            
            def askSheetHandle(k) = { k(sheet) }

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


            def getCell(addr, k) = {
                let rx = SheetIO.Common.CellAddress.getRow(addr);
                let cx = SheetIO.Common.CellAddress.getColumn(addr);
                let jsheet1: JSheet = checked_cast(jsheet);
                let ans = unsafe IO as r { JSheetUtil.getCell(jsheet1, rx, cx) };
                if (Object.isNull(ans)) 
                    k(None)
                else {
                    let jxssfcell: JXSSFCell = unchecked_cast(ans as JXSSFCell);
                    k(Some(XssfCell.XssfCell(jxssfcell)))
                }
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
    
    
    pub def withRow(ix: Int32, f: Unit -> a \ ef): Result[String, a] \ (ef + Sheet) - Row = 
        match Sheet.getRow(ix) {
            case Some(row) => SheetIO.XSSF.Row.runRow(row, f)
            case None      => Err("No Row")
        }

   
    pub def withCell(addr: CellAddress, f: Unit -> a \ ef): Result[String, a] \ (ef + Sheet) - Cell = 
        match Sheet.getCell(addr) {
            case Some(cell) => SheetIO.XSSF.Cell.runCell(cell, f)
            case None       => Err("No Cell")
        }

    pub def getCellToString(addr: CellAddress): Option[String] \ Sheet  = 
        withCell(addr, _ -> Cell.toString()) |> Result.toOption 


    pub def getCellRawValue(addr: CellAddress): Option[String] \ Sheet  = 
        withCell(addr, _ -> Cell.getRawValue()) |> Result.toOption |> Option.flatten

    pub def getStringCellValue(addr: CellAddress): Option[String] \ Sheet  = 
        withCell(addr, _ -> Cell.getStringCellValue()) |> Result.toOption |> Option.flatten

    pub def getFloat64CellValue(addr: CellAddress): Option[Float64] \ Sheet  = 
        withCell(addr, _ -> Cell.getFloat64CellValue()) |> Result.toOption |> Option.flatten


    pub def rowIterator(rc: Region[r]): Iterator[XssfRow, r, r] \ Sheet + r = 
        let XssfSheet.XssfSheet(jsheet) = Sheet.askSheetHandle();
        let javaIter: JIterator[JRow] = unsafe IO as r { jsheet.rowIterator() };
        def wrap(jrow) = { 
            let jxssfrow = unchecked_cast(jrow as JXSSFRow);
            XssfRow.XssfRow(jxssfrow)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JRow]), javaIter)
            |> Iterator.map(wrap)
    
    pub def withEachRow(f: Unit -> Unit \ ef): Unit \ Sheet + ef - Row = region rc {
        Iterator.forEach(row -> discard SheetIO.XSSF.Row.runRow(row, f),  rowIterator(rc))
    }

    pub def foldRows(f: st -> st \ ef, state: st): Result[String, st] \ Sheet + ef - Row = region rc {
        let iter = rowIterator(rc);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None => Ok(st)
            case Some(row) => match SheetIO.XSSF.Row.runRow(row, _ -> f(st)) {
                case Err(msg) => Err(msg)
                case Ok(st1)  => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapRows(f: Unit -> Option[a] \ ef): Result[String, List[a]] \ Sheet + ef - Row = region rc {
        let iter = rowIterator(rc);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None => Ok(k(Nil))
            case Some(row) => match SheetIO.XSSF.Row.runRow(row, _ -> f()) {
                case Err(msg) => Err(msg)
                case Ok(None)  => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


}

