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


pub mod SheetIO.Usermodel.Sheet {
    use ToJava.toJava
    
    use EffectExtras.Alt.Alt

    use SheetIO.Common.CellAddress.CellAddress
    use SheetIO.Usermodel.CellReader.CellReader
    use SheetIO.Usermodel.CellReader.SSCell
    use SheetIO.Usermodel.Row.Row    
    use SheetIO.Usermodel.Row.SSRow
    use SheetIO.Usermodel.Workbook.SSWorkbook


    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import org.apache.poi.ss.util.{SheetUtil => JSheetUtil}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}


    pub eff Sheet {
        // `askSheetHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askSheetHandle(): SSSheet
        def getSheetName() : String
        def getRow(i: Int32): SSRow
        def getCell(addr: CellAddress): SSCell
        def getFirstRowNum(): Int32
        def getLastRowNum(): Int32
        def getProtect(): Bool
        def getWorkbook(): SSWorkbook
    }

    pub enum SSSheet(JSheet)

    instance ToJava[SSSheet] {
        type Out = JSheet
        pub def toJava(x: SSSheet): JSheet = 
            let SSSheet.SSSheet(x1) = x;
            x1
    }

    instance ToFlix[SSSheet] {
        type In = JSheet
        pub def toFlix(x: JSheet): SSSheet = SSSheet.SSSheet(x)
    }

    pub enum CellMissingPolicy with Eq, Order, ToString {
        case CREATE_NULL_AS_BLANK 
        case RETURN_BLANK_AS_NULL 
        case RETURN_NULL_AND_BLANK 
    }
    
    pub def handle(sheet: SSSheet, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - Sheet) + r = x -> {
        let SSSheet.SSSheet(jsheet) = sheet;
        run {
            f(x)
        } with handler Sheet {
            
            def askSheetHandle(k) = { k(sheet) }

            def getSheetName(k) = try {
                let ans = unsafe IO as {ef + Alt - Sheet} { jsheet.getSheetName() };
                k(ans)
            } catch {
                case ex: JException => Alt.failure(ex.getMessage())
            }


            def getRow(i, k) = {
                let ans = unsafe IO as {ef + Alt - Sheet} { jsheet.getRow(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getRow")
            } 


            def getCell(addr, k) = {
                let rx = SheetIO.Common.CellAddress.getRow(addr);
                let cx = SheetIO.Common.CellAddress.getColumn(addr);
                let ans = unsafe IO as {ef + Alt - Sheet} { JSheetUtil.getCell(jsheet, rx, cx) };
                if (not Object.isNull(ans)) {
                    k(ToFlix.toFlix(ans))
                } else {
                    Alt.failure("getCell ${addr}")
                }
            }

            def getFirstRowNum(k) = {
                let ans = unsafe IO as {ef + Alt - Sheet} { jsheet.getFirstRowNum() };
                k(ans)
            } 

            def getLastRowNum(k) = {
                let ans = unsafe IO as {ef + Alt - Sheet} { jsheet.getLastRowNum() };
                k(ans)
            } 
            
            def getProtect(k) = {
                let ans = unsafe IO as {ef + Alt - Sheet} { jsheet.getProtect() };
                k(ans)
            }

            def getWorkbook(k) = {
                let ans = unsafe IO as {ef + Alt - Sheet} { jsheet.getWorkbook() };
                k(ToFlix.toFlix(ans))
            } 
        }
    }

    pub def runSheet(sheet: SSSheet, f: Unit -> a \ ef): a \ (ef + Alt - Sheet) = 
        handle(sheet, f)()
    
    pub def runSheetWithResult(sheet: SSSheet, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - Sheet) = 
        run { handle(sheet, f)() } with EffectExtras.Alt.runWithResult
    
    pub def withRow(ix: Int32, f: Unit -> a \ ef): a \ (ef + Alt + Sheet) - Row = 
        let row = Sheet.getRow(ix);
        SheetIO.Usermodel.Row.runRow(row, f)

   
    pub def withCell(addr: CellAddress, f: Unit -> a \ ef): Result[String, a] \ (ef + Sheet) - CellReader - Alt = 
        let cell = Sheet.getCell(addr);
        SheetIO.Usermodel.CellReader.runCellReaderWithResult(cell, f)

    pub def getCellToString(addr: CellAddress): Option[String] \ Sheet  = 
        withCell(addr, _ -> CellReader.toString()) |> Result.toOption 


    pub def getStringCellValue(addr: CellAddress): Option[String] \ Sheet  = 
        withCell(addr, _ -> CellReader.getStringCellValue()) |> Result.toOption 

    pub def getFloat64CellValue(addr: CellAddress): Option[Float64] \ Sheet  = 
        withCell(addr, _ -> CellReader.getFloat64CellValue()) |> Result.toOption


    pub def rowIterator(rc: Region[r]): Iterator[SSRow, r, r] \ Sheet + r = 
        let jsheet = Sheet.askSheetHandle() |> ToJava.toJava;
        let javaIter: JIterator[JRow] = unsafe IO as r { jsheet.rowIterator() };
        def wrap(jrow) = { 
            ToFlix.toFlix(jrow)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JRow]), javaIter)
            |> Iterator.map(wrap)
    
    pub def withEachRow(f: Unit -> Unit \ ef): Unit \ Sheet + ef - Alt - Row = region rc {
        Iterator.forEach(row -> discard SheetIO.Usermodel.Row.runRowWithResult(row, f),  rowIterator(rc))
    }

    pub def foldRows(f: st -> st \ ef, state: st): Result[String, st] \ Sheet + ef - Row = region rc {
        let iter = rowIterator(rc);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None => Ok(st)
            case Some(row) => match SheetIO.Usermodel.Row.runRowWithResult(row, _ -> f(st)) {
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
            case Some(row) => match SheetIO.Usermodel.Row.runRowWithResult(row, _ -> f()) {
                case Err(msg) => Err(msg)
                case Ok(None)  => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


}

