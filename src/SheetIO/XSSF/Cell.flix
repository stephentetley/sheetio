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


pub mod SheetIO.XSSF.Cell {

    use EffectExtras.Alt.Alt
    use SheetIO.XSSF.Comment.XssfComment

    import org.apache.poi.xssf.usermodel.{XSSFCell => JXSSFCell}
    import java.lang.{Exception => JException}

    pub enum XssfCell(JXSSFCell)

    // Make the "read" functions return Option[_] so we can do an alternative

    pub eff Cell {
        def getRawValue(): String
        def getBoolCellValue(): Bool
        def getFloat64CellValue(): Float64
        def getStringCellValue(): String
        def getCellFormula(): String
        def getCellComment(): XssfComment
        def getColumnIndex(): Int32        
        def getReference(): String
        def getRowIndex(): Int32
        def toString(): String
    }

    pub def handle(cell: XssfCell, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - Cell) = x -> 
        let XssfCell.XssfCell(jcell) = cell;
        run {
            f(x)
        } with handler Cell {

            def getRawValue(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getRawValue() };
                if (not Object.isNull(ans)) 
                    k(ans)
                else 
                    Alt.failure("getRawValue")
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
                    k(XssfComment.XssfComment(ans))
                else 
                     Alt.failure("getCellComment")
                    
            } 

            def getColumnIndex(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getColumnIndex() };
                k(ans)
            } 


            def getReference(k) = {
                let ans = unsafe IO as {ef + Alt - Cell} { jcell.getReference() };
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


    pub def runCell(cell: XssfCell, f: Unit -> a \ ef): a \ (ef + Alt - Cell) = 
        handle(cell, f)()

    pub def runCellWithResult(cell: XssfCell, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - Cell) = 
        run { handle(cell, f)() } with EffectExtras.Alt.runWithResult

}