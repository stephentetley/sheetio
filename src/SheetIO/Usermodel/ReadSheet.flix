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

    use SheetIO.Usermodel.ReadRow.ReadRow
    use SheetIO.Usermodel.ReadCell.ReadCell
    use SheetIO.Usermodel.{Workbook, Sheet, Row, Cell}

    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import org.apache.poi.ss.util.{SheetUtil => JSheetUtil}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}


    pub eff ReadSheet {
        // `askSheetHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askSheetHandle(): Sheet
        def getSheetName() : String
        def getRow(i: Int32): Row
        def getCell(addr: CellAddress): Cell
        def getFirstRowNum(): Int32
        def getLastRowNum(): Int32
        def getProtect(): Bool
        def getWorkbook(): Workbook
    }


    pub enum CellMissingPolicy with Eq, Order, ToString {
        case CREATE_NULL_AS_BLANK 
        case RETURN_BLANK_AS_NULL 
        case RETURN_NULL_AND_BLANK 
    }
    
    pub def handle(sheet: Sheet,
                    f: a -> b \ ef): a -> b \ (ef + Fail - ReadSheet) + r = x -> {
        let jsheet = ToJava.toJava(sheet);
        run {
            f(x)
        } with handler ReadSheet {
            
            def askSheetHandle(k) = { k(sheet) }

            def getSheetName(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadSheet} { jsheet.getSheetName() };
                k(ans)
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }


            def getRow(i, k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet} { jsheet.getRow(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Fail.report("getRow ${i}")
            } 


            def getCell(addr, k) = {
                let rx = SheetIO.Common.CellAddress.getRow(addr);
                let cx = SheetIO.Common.CellAddress.getColumn(addr);
                let ans = unsafe IO as {ef + Fail - ReadSheet} { JSheetUtil.getCell(jsheet, rx, cx) };
                if (not Object.isNull(ans)) {
                    k(ToFlix.toFlix(ans))
                } else {
                    Fail.report("getCell ${addr}")
                }
            }

            def getFirstRowNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet} { jsheet.getFirstRowNum() };
                k(ans)
            } 

            def getLastRowNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet} { jsheet.getLastRowNum() };
                k(ans)
            } 
            
            def getProtect(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet} { jsheet.getProtect() };
                k(ans)
            }

            def getWorkbook(k) = {
                let ans = unsafe IO as {ef + Fail - ReadSheet} { jsheet.getWorkbook() };
                k(ToFlix.toFlix(ans))
            } 
        }
    }

    pub def runReadSheet(sheet: Sheet, f: Unit -> a \ ef): a \ (ef + Fail - ReadSheet) =
        handle(sheet, f)()
    

    pub def runReadSheetWithResult(sheet: Sheet, f: Unit -> a \ ef): Result[String, a] \ (ef - Fail - ReadSheet) =
        run { handle(sheet, f)() } with EffectExtras.Fail.runWithResult


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


    pub def rowIteratorX(rc: Region[r], sheet: Sheet): Iterator[Row, r, r] \ r =
        let jsheet = ToJava.toJava(sheet);
        let javaIter: JIterator[JRow] = unsafe IO as r { jsheet.rowIterator() };
        def wrap(jrow) = { 
            ToFlix.toFlix(jrow)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JRow]), javaIter)
            |> Iterator.map(wrap)

    pub def rowIterator(rc: Region[r]): Iterator[Row, r, r] \ ReadSheet + r =
        let jsheet = ReadSheet.askSheetHandle() |> ToJava.toJava;
        let javaIter: JIterator[JRow] = unsafe IO as r { jsheet.rowIterator() };
        def wrap(jrow) = { 
            ToFlix.toFlix(jrow)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JRow]), javaIter)
            |> Iterator.map(wrap)
    
    pub def forEachRow(f: Unit -> Unit \ ef): Unit \ ReadSheet + ef - Fail - ReadRow = region rc {
        Iterator.forEach(row -> discard SheetIO.Usermodel.ReadRow.runReadRowWithResult(row, f),  rowIterator(rc))
    }

    pub def foldRows(f: st -> st \ ef, state: st): Result[String, st] \ ReadSheet + ef - ReadRow = region rc {
        let iter = rowIterator(rc);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None => Ok(st)
            case Some(row) => match SheetIO.Usermodel.ReadRow.runReadRowWithResult(row, _ -> f(st)) {
                case Err(msg) => Err(msg)
                case Ok(st1)  => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapRows(f: Unit -> Option[a] \ ef): Result[String, List[a]] \ ef + ReadSheet - ReadRow = region rc {
        let iter = rowIterator(rc);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None => Ok(k(Nil))
            case Some(row) => match SheetIO.Usermodel.ReadRow.runReadRowWithResult(row, _ -> f()) {
                case Err(msg) => Err(msg)
                case Ok(None)  => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


}

