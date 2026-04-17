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


pub mod SheetIO.XSSF.Row {

    use IoError.ErrorKind
    use IoError.IoError

    use SheetIO.XSSF.Cell.XssfCell

    import org.apache.poi.xssf.usermodel.{XSSFRow => JXSSFRow}
    import java.lang.{Exception => JException}

    pub enum XssfRow(JXSSFRow)

    pub eff Row {
        def getCell(i: Int32): XssfCell
        def getFirstCellNum(): Option[Int32]
        def getLastCellNum(): Option[Int32]
        def getRowNum() : Int32
    }

    pub def handle(row: XssfRow, 
                    f: a -> b \ ef): a -> Result[String, b] \ ef - Row = x -> region rc {
        handleHelper(rc, row, f, x)
    }

    def handleHelper(_: Region[r],
                        row: XssfRow, 
                        f: a -> b \ ef): a -> Result[String, b] \ (ef - Row) + r  = x -> {
        let XssfRow.XssfRow(jrow) = row;
        run {
            Ok(f(x))
        } with handler Row {
            
            def getCell(i, k) = {
                let ans = unchecked_cast(jrow.getCell(i) as _ \ r);
                k(XssfCell.XssfCell(ans))
            }

            def getFirstCellNum(k) = {
                let ans = unchecked_cast(jrow.getFirstCellNum() as _ \ r);
                if (ans < 0i16) 
                    k(None)
                else
                    k(Some(Int16.toInt32(ans)))
            }

            def getLastCellNum(k) = {
                let ans = unchecked_cast(jrow.getLastCellNum() as _ \ r);
                if (ans < 0i16) 
                    k(None)
                else
                    k(Some(Int16.toInt32(ans)))
            }

            def getRowNum(k) = {
                let ans = unchecked_cast(jrow.getRowNum() as _ \ r);
                k(ans)
            } 
        }
    }

    pub def runRow(row: XssfRow, f: Unit -> a \ ef): Result[String, a] \ (ef - Row) = 
        handle(row, f)()

}