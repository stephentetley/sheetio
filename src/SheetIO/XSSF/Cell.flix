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

    use IoError.ErrorKind
    use IoError.IoError

    import org.apache.poi.xssf.usermodel.{XSSFCell => JXSSFCell}
    import java.lang.{Exception => JException}

    pub enum XssfCell(JXSSFCell)

    pub eff Cell {
        def getColumnIndex(): Int32
        def getRawValue(): Option[String]
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
            
            def getColumnIndex(k) = {
                let ans = unchecked_cast(jcell.getColumnIndex() as _ \ r);
                k(ans)
            } 


            def getRawValue(k) = {
                let ans = unchecked_cast(jcell.getRawValue() as _ \ r);
                if (Object.isNull(ans)) 
                    k(None)
                else 
                    k(Some(ans))
            } 

            def getReference(k) = {
                let ans = unchecked_cast(jcell.getReference() as _ \ r);
                k(ans)
            } 

            def getRowIndex(k) = {
                let ans = unchecked_cast(jcell.getRowIndex() as _ \ r);
                k(ans)
            } 

            def toString(k) = {
                let ans = unchecked_cast(jcell.toString() as _ \ r);
                k(ans)
            } 
        }
    }

    pub def runCell(cell: XssfCell, f: Unit -> a \ ef): Result[String, a] \ (ef - Cell) = 
        handle(cell, f)()

}