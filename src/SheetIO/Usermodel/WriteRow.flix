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


pub mod SheetIO.Usermodel.WriteRow {

    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail


    use SheetIO.Usermodel.{Sheet, Row, Cell, CellStyle}

    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell} // superclass
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}

    pub eff WriteRow {
        def createCell(i: Int32): Cell
        def removeCell(c: Cell): Unit
        def setRowStyle(style: CellStyle): Unit
    }

        
    pub def handle(row: Row, 
                    f: a -> b \ ef): a -> b \ (ef + Fail - WriteRow) + r  = x -> {
        let jrow = ToJava.toJava(row);
        run {
            f(x)
        } with handler WriteRow {

            def createCell(i, k) = {
                let ans = unsafe IO as {ef + Fail - WriteRow} { jrow.createCell(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Fail.report("createCell ${i}")
            }
            
            def removeCell(cell, k) = {
                let cell1 = ToJava.toJava(cell);
                unsafe IO as {ef + Fail - WriteRow} { jrow.removeCell(cell1) };
                k()
            }
            
            def setRowStyle(style, k) = {
                let style1 = ToJava.toJava(style);
                unsafe IO as {ef + Fail - WriteRow} { jrow.setRowStyle(style1) };
                k()
            }

        }
    }

    pub def runWriteRow(row: Row, f: Unit -> a \ ef): a \ (ef + Fail - WriteRow) = 
        handle(row, f)()

    pub def runWriteRowWithResult(row: Row, 
                                    f: Unit -> a \ ef): Result[String, a] \ (ef - Fail - WriteRow) = 
        run { handle(row, f)() } with EffectExtras.Fail.runWithResult



}

