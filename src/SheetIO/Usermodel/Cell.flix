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


pub mod SheetIO.Usermodel.Cell {
    
    
    use EffectExtras.Alt.Alt
    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.Comment.SSComment

    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import java.lang.{Exception => JException}


    pub eff Cell {
        def getAddress(): CellAddress
        def getCellType(): CellType
        def getBoolCellValue(): Bool
        def getFloat64CellValue(): Float64
        def getStringCellValue(): String
        def getCellFormula(): String
        def getCellComment(): SSComment
        def getColumnIndex(): Int32        
        def getRowIndex(): Int32
        def toString(): String
    }

    pub enum SSCell(JCell)

    instance ToJava[SSCell] {
        type Out = JCell
        pub def toJava(x: SSCell): JCell = 
            let SSCell.SSCell(x1) = x;
            x1
    }

    instance ToFlix[SSCell] {
        type In = JCell
        pub def toFlix(x: JCell): SSCell = SSCell.SSCell(x)
    }

    pub def handle(cell: SSCell, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - Cell) = x -> 
        let jcell = ToJava.toJava(cell);
        run {
            f(x)
        } with handler Cell {

            def getAddress(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getAddress() };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCellAddress")
            }

            def getCellType(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getCellType() };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCellType")
            }

            def getBoolCellValue(k) = try {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getBooleanCellValue() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getBoolCellValue")
            }

            def getFloat64CellValue(k) = try {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getNumericCellValue() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getFloat64CellValue")
            }

            def getStringCellValue(k) = try {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getStringCellValue() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getStringCellValue")
            }

            def getCellFormula(k) = try {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getCellFormula() };
                k(ans)
            } catch {
                case _: JException => Alt.failure("getCellFormula")
            }

            def getCellComment(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getCellComment() };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getCellComment")
            }

            def getColumnIndex(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getColumnIndex() };
                k(ans)
            } 

            def getRowIndex(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getRowIndex() };
                k(ans)
            } 

            def toString(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.toString() };
                k(ans)
            } 

        }


    pub def runCell(cell: SSCell, f: Unit -> a \ ef): a \ (ef + Alt - Cell) = 
        handle(cell, f)()

    pub def runCellWithResult(cell: SSCell, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - Cell) = 
        run { handle(cell, f)() } with EffectExtras.Alt.runWithResult

}
