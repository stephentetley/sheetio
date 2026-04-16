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

    use SheetIO.XSSF.Sheet.XssfSheet

    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import java.lang.{Exception => JException}
    import java.io.{FileInputStream => JFileInputStream}
    import java.io.{File => JFile}

    pub enum XssfWorkbook(JXSSFWorkbook)

    pub eff Workbook {
        def createSheet(): XssfSheet
        def createSheetWithName(s: String): XssfSheet
        def getNumberOfSheets(): Int32
        def getSheetAt(i: Int32): XssfSheet
        def getSheetIndex(s: String): Option[Int32]
        def getSheetIndexOfSheet(s: XssfSheet): Option[Int32]

    }


    pub def handle(wb: XssfWorkbook, 
                    f: a -> b \ ef): a -> Result[IoError, b] \ {IO + ef} - Workbook = x -> {
        let XssfWorkbook.XssfWorkbook(jworkbook) = wb;
        run {
            Ok(f(x))
        } with handler Workbook {
            
            def createSheet(k) = {
                let ans = jworkbook.createSheet();
                k(XssfSheet.XssfSheet(ans))
            }
           
            def createSheetWithName(s, k) = {
                let ans = jworkbook.createSheet(s);
                k(XssfSheet.XssfSheet(ans))
            }

            def getNumberOfSheets(k) = try {
                let ans = jworkbook.getNumberOfSheets();
                k(ans)
            } catch {
                case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
            }

            def getSheetAt(i, k) = try {
                let ans = jworkbook.getSheetAt(i);
                k(XssfSheet.XssfSheet(ans))
            } catch {
                case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
            }

            def getSheetIndex(s, k) = {
                let ans = jworkbook.getSheetIndex(s);
                if (ans < 0) k(None) else k(Some(ans))
            }

            def getSheetIndexOfSheet(s, k) = {
                let XssfSheet.XssfSheet(jsheet) = s;
                let ans = jworkbook.getSheetIndex(jsheet);
                if (ans < 0) k(None) else k(Some(ans))
            }

        }
    }

    pub def runReadOnly(path: String, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Workbook = try {
        let jfile = new JFile(path);
        let jfis = new JFileInputStream(jfile);        
        let jworkbook = new JXSSFWorkbook(jfis);
        let ans = handle(XssfWorkbook.XssfWorkbook(jworkbook), f)();
        jworkbook.close();
        jfis.close();
        ans
    } catch {
        case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
    }

    pub def runReadWrite(path: String, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Workbook = try {
        let jworkbook = new JXSSFWorkbook(path);
        let ans = handle(XssfWorkbook.XssfWorkbook(jworkbook), f)();
        jworkbook.close();
        ans
    } catch {
        case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
    }

}