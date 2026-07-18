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


pub mod SheetIO.Usermodel.Row {

    use EffectExtras.Alt.Alt

    use SheetIO.Usermodel.Sheet.SSSheet
    use SheetIO.Usermodel.ReadCell.ReadCell
    use SheetIO.Usermodel.{Cell}

    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell} // superclass
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}

    pub eff Row {
        // `askRowHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askRowHandle(): SSRow       
        def getCell(i: Int32): Cell
        def getFirstCellNum(): Int32
        def getLastCellNum(): Int32
        def getRowNum(): Int32
        def getSheet(): SSSheet
    }

    pub enum SSRow(JRow)
    
    instance ToJava[SSRow] {
        type Out = JRow
        pub def toJava(x: SSRow): JRow = 
            let SSRow.SSRow(x1) = x;
            x1
    }

    instance ToFlix[SSRow] {
        type In = JRow
        pub def toFlix(x: JRow): SSRow = SSRow.SSRow(x)
    }
        
    pub def handle(row: SSRow, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - Row) + r  = x -> {
        let SSRow.SSRow(jrow) = row;
        run {
            f(x)
        } with handler Row {
            
            def askRowHandle(k) = k(row)

            def getCell(i, k) = {
                let ans = unsafe IO as {ef + Alt - Row} { jrow.getCell(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCell ${i}")
            }

            def getFirstCellNum(k) = {
                let ans = unsafe IO as {ef + Alt - Row} { jrow.getFirstCellNum() };
                if (ans >= 0i16) 
                    k(Int16.toInt32(ans))
                else
                    Alt.failure("getFirstCellNum")
            }

            def getLastCellNum(k) = {
                let ans = unsafe IO as {ef + Alt - Row} { jrow.getLastCellNum() };
                if (ans >= 0i16) 
                    k(Int16.toInt32(ans))
                else
                    Alt.failure("getLastCellNum")
            }

            def getRowNum(k) = {
                let ans = unsafe IO as {ef + Alt - Row} { jrow.getRowNum() };
                k(ans)
            } 

            def getSheet(k) = {
                let ans = unsafe IO as {ef + Alt - Row} { jrow.getSheet() };
                k(SSSheet.SSSheet(ans))
            } 

        }
    }

    pub def runRow(row: SSRow, f: Unit -> a \ ef): a \ (ef + Alt - Row) = 
        handle(row, f)()

    pub def runRowWithResult(row: SSRow, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - Row) = 
        run { handle(row, f)() } with EffectExtras.Alt.runWithResult


    pub def withCell(ix: Int32,f: Unit -> a \ ef): a \ (ef + Alt + Row) - ReadCell = 
        let cell =  Row.getCell(ix);
        SheetIO.Usermodel.ReadCell.runReadCell(cell, f)
    
    pub def cellIterator(rc: Region[r]): Iterator[Cell, r, r] \ Row + r = 
        let jrow = Row.askRowHandle() |> ToJava.toJava;
        let javaIter: JIterator[JCell] = unsafe IO as r { jrow.cellIterator() };
        def wrap(jcell) = { 
            ToFlix.toFlix(jcell)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JCell]), javaIter)
            |> Iterator.map(wrap)
    
    pub def withEachCell(f: Unit -> Unit \ ef): Unit \ Row + ef - ReadCell - Alt = region rc {
        Iterator.forEach(cell -> 
            discard SheetIO.Usermodel.ReadCell.runReadCellWithResult(cell, f),  cellIterator(rc))
    }

    pub def foldCells(f: st -> st \ ef, state: st): Result[String, st] \ Row + ef - ReadCell = region rc {
        let iter = cellIterator(rc);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None => Ok(st)
            case Some(cell) => match SheetIO.Usermodel.ReadCell.runReadCellWithResult(cell, _ -> f(st)) {
                case Err(msg) => Err(msg)
                case Ok(st1)  => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapCells(f: Unit -> Option[a] \ ef): Result[String, List[a]] \ Row + ef - ReadCell = region rc {
        let iter = cellIterator(rc);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None => Ok(k(Nil))
            case Some(cell) => match SheetIO.Usermodel.ReadCell.runReadCellWithResult(cell, _ -> f()) {
                case Err(msg) => Err(msg)
                case Ok(None)  => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


}

