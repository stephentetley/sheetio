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


pub mod SheetIO.Usermodel.ReadSheet {


    
    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail

    use SheetIO.Common.CellAddress.CellAddress
    use SheetIO.Common.CellAddress.ToCellAddress

    use SheetIO.Usermodel.ReadWorkbook.ReadWorkbook
    use SheetIO.Usermodel.ReadRow.ReadRow
    use SheetIO.Usermodel.ReadCell.ReadCell
    use SheetIO.Usermodel.{Workbook, Sheet, Row, Cell}
    use SheetIO.Usermodel.ReifySheet
    use SheetIO.Usermodel.ReifyCell

    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import org.apache.poi.ss.util.{SheetUtil => JSheetUtil}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}




    pub eff ReadSheet {
        def getSheetName() : String
        def getRow(i: Int32): Row
        def getCell(addr: CellAddress): Cell
        def getFirstRowNum(): Int32
        def getLastRowNum(): Int32
        def getProtect(): Bool
        def getWorkbook(): Workbook
        def isColumnHidden(ix: Int32): Bool
    }


    pub enum CellMissingPolicy with Eq, Order, ToString {
        case CREATE_NULL_AS_BLANK 
        case RETURN_BLANK_AS_NULL 
        case RETURN_NULL_AND_BLANK 
    }
    
    pub def handle(sheet: Sheet,
                    f: a -> b \ ef): a -> b \ (ef + Fail - ReadSheet - ReifySheet) = x -> {
        let jsheet = ToJava.toJava(sheet);
        run {
            f(x)
        } with handler ReifySheet {
            
            def getSheet(k) = k(sheet)

        } with handler ReadSheet {

            def getSheetName(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.getSheetName() };
                k(ans)
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }


            def getRow(i, k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.getRow(i) };
                EffectExtras.Fail.assertNotNullWithMessage("getRow ${i}", ans);
                k(ToFlix.toFlix(ans))
            } 


            def getCell(addr, k) = {
                let rx = SheetIO.Common.CellAddress.getRow(addr);
                let cx = SheetIO.Common.CellAddress.getColumn(addr);
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { JSheetUtil.getCell(jsheet, rx, cx) };
                EffectExtras.Fail.assertNotNullWithMessage("getCell ${addr}", ans);
                k(ToFlix.toFlix(ans))
            }

            def getFirstRowNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.getFirstRowNum() };
                k(ans)
            } 

            def getLastRowNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.getLastRowNum() };
                k(ans)
            } 
            
            def getProtect(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.getProtect() };
                k(ans)
            }

            def getWorkbook(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.getWorkbook() };
                k(ToFlix.toFlix(ans))
            } 

            def isColumnHidden(ix, k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet - ReifySheet} { jsheet.isColumnHidden(ix) };
                k(ans)
            }
        }
    }

    pub def runReadSheet(sheet: Sheet, f: Unit -> a \ ef): a \ (ef + Fail - ReadSheet - ReifySheet) =
        handle(sheet, f)()
 


    // ReadSheet context
    pub def readRowAt(ix: Int32, f: Unit -> a \ ef): a \ ef + Fail + ReadSheet - ReadRow =
        let row = ReadSheet.getRow(ix);
        SheetIO.Usermodel.ReadRow.runReadRow(row, f)


    // ReadSheet context
    pub def readCellAddr(cellAddr: addr, f: Unit -> a \ ef): a \ ef + Fail + ReadSheet - ReadCell - ReifyCell - Alt with ToCellAddress[addr] = 
        let addr1 = ToCellAddress.toCellAddress(cellAddr) |> Fail.fromOption;
        let cell = ReadSheet.getCell(addr1);
        SheetIO.Usermodel.ReadCell.runReadCell(cell, f)


    pub def getCellValueBool(cellAddr: addr): Bool \ ReadSheet + Fail with ToCellAddress[addr] = 
        let addr1 = ToCellAddress.toCellAddress(cellAddr) |> Fail.fromOption;
        let cell = ReadSheet.getCell(addr1);
        SheetIO.Usermodel.ReadCell.runReadCell(cell, ReadCell.getCellValueBool)

    pub def getCellValueFloat64(cellAddr: addr): Float64 \ ReadSheet + Fail with ToCellAddress[addr] = 
        let addr1 = ToCellAddress.toCellAddress(cellAddr) |> Fail.fromOption;
        let cell = ReadSheet.getCell(addr1);
        SheetIO.Usermodel.ReadCell.runReadCell(cell, ReadCell.getCellValueFloat64)
    
    pub def getCellValueString(cellAddr: addr): String \ ReadSheet + Fail with ToCellAddress[addr] = 
        let addr1 = ToCellAddress.toCellAddress(cellAddr) |> Fail.fromOption;
        let cell = ReadSheet.getCell(addr1);
        SheetIO.Usermodel.ReadCell.runReadCell(cell, ReadCell.getCellValueString)


    pub def rowIterator(rc: Region[r], sheet: Sheet): Iterator[Row, r, r] \ r =
        let jsheet = ToJava.toJava(sheet);
        let javaIter: JIterator[JRow] = unsafe IO as r { jsheet.rowIterator() };
        def wrap(jrow) = { 
            ToFlix.toFlix(jrow)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JRow]), javaIter)
            |> Iterator.map(wrap)

    pub def forEachRow(f: Row -> Unit \ ef): Unit \ ef + ReifySheet = region rc {
        let sheet = ReifySheet.getSheet();
        Iterator.forEach(f, rowIterator(rc, sheet))
    }


    pub def foldRows(f: Row -> st -> st \ ef, state: st): st \ ef + Fail + ReifySheet = region rc {
        let sheet = ReifySheet.getSheet();
        let iter = rowIterator(rc, sheet);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None       => st
            case Some(row) => match EffectExtras.Fail.runWithResult(_ -> f(row, st)) {
                case Err(msg)   => Fail.report(msg)
                case Ok(st1)    => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapRows(f: Row -> Option[a] \ ef): List[a] \ ef + Fail + ReifySheet = region rc {
        let sheet = ReifySheet.getSheet();
        let iter = rowIterator(rc, sheet);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None       => k(Nil)
            case Some(row) => match EffectExtras.Fail.runWithResult(_ -> f(row)) {
                case Err(msg)       => Fail.report(msg)
                case Ok(None)       => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


}

