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


pub mod SheetIO.Workbook {

    use IoError.ErrorKind
    use IoError.IoError

    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import java.lang.{Exception => JException}

    // pub enum WorkbookHandle(JXSSFWorkbook)

    pub eff Workbook {
        def getNumberOfSheets() : Int32
    }

    pub def handle(path: String, 
                f: a -> b \ ef): a -> Result[IoError, b] \ {IO + ef} - Workbook = x -> {
        let jworkbook = new JXSSFWorkbook(path);
        let ans = run {
            Ok(f(x))
        } with handler Workbook {
            
            def getNumberOfSheets(k) = try {
                let ans = jworkbook.getNumberOfSheets();
                k(ans)
            } catch {
                case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
            }
        };
        jworkbook.close();
        ans
    }

    pub def runWithIO(path: String, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Workbook = 
        handle(path, f)()

}