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


pub mod SheetIO.Usermodel.ReadRow {
    use SheetIO.Usermodel.ReadSheet

    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail

    use SheetIO.Usermodel.ReadCell.ReadCell
    use SheetIO.Usermodel.{Sheet, Row, Cell}
    use SheetIO.Usermodel.ReifyRow

    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell} // superclass
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}

    pub eff ReadRow {
        def getCell(i: Int32): Cell
        def getFirstCellNum(): Int32
        def getLastCellNum(): Int32
        def getRowNum(): Int32
        def getSheet(): Sheet
    }

        
    pub def handle(row: Row, 
                    f: a -> b \ ef): a -> b \ (ef + Fail - ReadRow - ReifyRow) + r  = x -> {
        let jrow = ToJava.toJava(row);
        run {
            f(x)
        } with handler ReifyRow {
            
            def getRow(k) = k(row)

        } with handler ReadRow {
            
            def getCell(i, k) = {
                let ans = unsafe IO as {ef + Fail - ReadRow - ReifyRow} { jrow.getCell(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Fail.report("getCell ${i}")
            }

            def getFirstCellNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadRow - ReifyRow} { jrow.getFirstCellNum() };
                if (ans >= 0i16) 
                    k(Int16.toInt32(ans))
                else
                    Fail.report("getFirstCellNum")
            }

            def getLastCellNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadRow - ReifyRow} { jrow.getLastCellNum() };
                if (ans >= 0i16) 
                    k(Int16.toInt32(ans))
                else
                    Fail.report("getLastCellNum")
            }

            def getRowNum(k) = {
                let ans = unsafe IO as {ef + Fail - ReadRow - ReifyRow} { jrow.getRowNum() };
                k(ans)
            } 

            def getSheet(k) = {
                let ans = unsafe IO as {ef + Fail - ReadRow - ReifyRow} { jrow.getSheet() };
                k(ToFlix.toFlix(ans))
            } 

        }
    }

    pub def runReadRow(row: Row, f: Unit -> a \ ef): a \ (ef + Fail - ReadRow) = 
        handle(row, f)()
    



    // ReadRow context
    pub def readCellAt(ix: Int32, f: Unit -> a \ ef): a \ ef + Fail + ReadRow - ReadCell - Alt =
        let cell = ReadRow.getCell(ix);
        SheetIO.Usermodel.ReadCell.runReadCell(cell, f)

    pub def cellIterator(rc: Region[r], row: Row): Iterator[Cell, r, r] \ r = 
        let jrow = ToJava.toJava(row);
        let javaIter: JIterator[JCell] = unsafe IO as r { jrow.cellIterator() };
        def wrap(jcell) = { 
            ToFlix.toFlix(jcell)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JCell]), javaIter)
            |> Iterator.map(wrap)
    
    
    pub def forEachCell(f: Cell -> Unit \ ef, t: Row): Unit \ ef = region rc {
        Iterator.forEach(f, cellIterator(rc, t))
    }


    pub def foldCells(f: Cell -> st -> st \ ef, state: st): st \ ef + Fail + ReifyRow = region rc {
        let row = ReifyRow.getRow();
        let iter = cellIterator(rc, row);
        // @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None       => st
            case Some(cell) => match EffectExtras.Fail.runWithResult(_ -> f(cell, st)) {
                case Err(msg)   => Fail.report(msg)
                case Ok(st1)    => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapCells(f: Cell -> Option[a] \ ef): List[a] \ ef + Fail + ReifyRow = region rc {
        let row = ReifyRow.getRow();
        let iter = cellIterator(rc, row);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None       => k(Nil)
            case Some(cell) => match EffectExtras.Fail.runWithResult(_ -> f(cell)) {
                case Err(msg)       => Fail.report(msg)
                case Ok(None)       => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


}

