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

    use SheetIO.XSSF.Comment.XssfComment

    import org.apache.poi.xssf.usermodel.{XSSFCell => JXSSFCell}
    import java.lang.{Exception => JException}

    pub enum XssfCell(JXSSFCell)

    // Make the "read" functions return Option[_] so we can do an alternative

    pub eff Cell {
        def getRawValue(): Option[String]
        def getBoolCellValue(): Option[Bool]
        def getFloat64CellValue(): Option[Float64]
        def getStringCellValue(): Option[String]
        def getCellFormula(): Option[String]
        def getCellComment(): Option[XssfComment]
        def getColumnIndex(): Int32        
        def getReference(): String
        def getRowIndex(): Int32
        def toString(): String
    }

    pub def handle(cell: XssfCell, 
                    f: a -> b \ ef): a -> Result[String, b] \ (ef - Cell) = x -> region rc {
        handleHelper(rc, cell, f, x)
    }

    def handleHelper(_: Region[r],
                        cell: XssfCell, 
                        f: a -> b \ ef): a -> Result[String, b] \ (ef - Cell) + r = x -> {
        let XssfCell.XssfCell(jcell) = cell;
        run {
            Ok(f(x))
        } with handler Cell {

            def getRawValue(k) = {
                let ans = unsafe IO as r { jcell.getRawValue() };
                if (Object.isNull(ans)) 
                    k(None)
                else 
                    k(Some(ans))
            } 

            def getBoolCellValue(k) = try {
                let ans = unsafe IO as r { jcell.getBooleanCellValue() };
                k(Some(ans))
            } catch {
                case _: JException => k(None)
            }

            def getFloat64CellValue(k) = try {
                let ans = unsafe IO as r { jcell.getNumericCellValue() };
                k(Some(ans))
            } catch {
                case _: JException => k(None)
            }

            def getStringCellValue(k) = try {
                let ans = unsafe IO as r { jcell.getStringCellValue() };
                k(Some(ans))
            } catch {
                case _: JException => k(None)
            }

            def getCellFormula(k) = try {
                let ans = unsafe IO as r { jcell.getCellFormula() };
                k(Some(ans))
            } catch {
                case _: JException => k(None)
            }

            def getCellComment(k) = {
                let ans = unsafe IO as r { jcell.getCellComment() };
                if (Object.isNull(ans)) 
                    k(None)
                else 
                    k(Some(XssfComment.XssfComment(ans)))
            } 

            def getColumnIndex(k) = {
                let ans = unsafe IO as r { jcell.getColumnIndex() };
                k(ans)
            } 


            def getReference(k) = {
                let ans = unsafe IO as r { jcell.getReference() };
                k(ans)
            } 

            def getRowIndex(k) = {
                let ans = unsafe IO as r { jcell.getRowIndex() };
                k(ans)
            } 

            def toString(k) = {
                let ans = unsafe IO as r { jcell.toString() };
                k(ans)
            } 
        }
    }

    pub def runCell(cell: XssfCell, f: Unit -> a \ ef): Result[String, a] \ (ef - Cell) = 
        handle(cell, f)()

}