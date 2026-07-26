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


pub mod SheetIO.Usermodel.ReadCell {
    
    
    use EffectExtras.Alt.Alt
    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.{Cell, Comment}

    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import java.lang.{Exception => JException}


    pub eff ReadCell {
        def getAddress(): CellAddress
        def getCellType(): CellType
        def getCellValueBool(): Bool
        def getCellValueFloat64(): Float64
        def getCellValueString(): String
        def getCellFormula(): String
        def getCellComment(): Comment
        def getColumnIndex(): Int32
        def getRowIndex(): Int32
        def toString(): String
    }

    pub def handle(cell: Cell, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - ReadCell) = x -> 
        let jcell = ToJava.toJava(cell);
        run {
            f(x)
        } with handler ReadCell {

            def getAddress(k) = {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getAddress() };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCellAddress")
            }

            def getCellType(k) = {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getCellType() };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCellType")
            }

            def getCellValueBool(k) = try {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getBooleanCellValue() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getBoolCellValue")
            }

            def getCellValueFloat64(k) = try {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getNumericCellValue() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getFloat64CellValue")
            }

            def getCellValueString(k) = try {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getStringCellValue() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getStringCellValue")
            }

            def getCellFormula(k) = try {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getCellFormula() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getCellFormula")
            }

            def getCellComment(k) = {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getCellComment() };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCellComment")
            }

            def getColumnIndex(k) = {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getColumnIndex() };
                k(ans)
            } 

            def getRowIndex(k) = {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.getRowIndex() };
                k(ans)
            } 

            def toString(k) = {
                let ans = unsafe IO as {ef + Alt - ReadCell} { jcell.toString() };
                k(ans)
            } 

        }


    pub def runReadCell(cell: Cell, f: Unit -> a \ ef): a \ (ef + Alt - ReadCell) = 
        handle(cell, f)()

    pub def runReadCellWithResult(cell: Cell, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - ReadCell) = 
        run { handle(cell, f)() } with EffectExtras.Alt.runWithResult

}
