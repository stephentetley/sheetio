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

    import org.apache.poi.xssf.usermodel.{XSSFRow => JXSSFRow}
    import java.lang.{Exception => JException}

    pub enum XssfRow(JXSSFRow)

    pub eff Row {
        def getRowNum() : Int32
    }

    pub def handle(row: XssfRow, 
                f: a -> b \ ef): a -> Result[IoError, b] \ {IO + ef} - Row = x -> {
        let XssfRow.XssfRow(jrow) = row;
        run {
            Ok(f(x))
        } with handler Row {
            
            def getRowNum(k) = try {
                let ans = jrow.getRowNum();
                k(ans)
            } catch {
                case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
            }
        }
    }

    pub def runWithIO(row: XssfRow, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Row = 
        handle(row, f)()

}