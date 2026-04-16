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


pub mod SheetIO.XSSF.Workbook {

    use IoError.ErrorKind
    use IoError.IoError

    use SheetIO.XSSF.Sheet.SheetHandle

    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import java.lang.{Exception => JException}
    import java.io.{FileInputStream => JFileInputStream}
    import java.io.{File => JFile}

    pub enum XSSFWorkbook(JXSSFWorkbook)

    pub eff Workbook {
        def createSheet(): SheetHandle
        def getNumberOfSheets(): Int32

    }


    pub def handle(wb: XSSFWorkbook, 
                    f: a -> b \ ef): a -> Result[IoError, b] \ {IO + ef} - Workbook = x -> {
        let XSSFWorkbook.XSSFWorkbook(jworkbook) = wb;
        run {
            Ok(f(x))
        } with handler Workbook {
            
            def createSheet(k) = {
                let ans = jworkbook.createSheet();
                k(SheetHandle.SheetHandle(ans))
            }
            
            def getNumberOfSheets(k) = try {
                let ans = jworkbook.getNumberOfSheets();
                k(ans)
            } catch {
                case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
            }
        }
    }

    pub def runReadOnly(path: String, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Workbook = try {
        let jfile = new JFile(path);
        let jfis = new JFileInputStream(jfile);        
        let jworkbook = new JXSSFWorkbook(jfis);
        let ans = handle(XSSFWorkbook.XSSFWorkbook(jworkbook), f)();
        jworkbook.close();
        jfis.close();
        ans
    } catch {
        case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
    }

    pub def runReadWrite(path: String, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Workbook = try {
        let jworkbook = new JXSSFWorkbook(path);
        let ans = handle(XSSFWorkbook.XSSFWorkbook(jworkbook), f)();
        jworkbook.close();
        ans
    } catch {
        case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
    }

}