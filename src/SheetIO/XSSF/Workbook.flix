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
    
    use SheetIO.XSSF.Sheet.Sheet
    use SheetIO.XSSF.Sheet.XssfSheet

    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import org.apache.poi.xssf.usermodel.{XSSFSheet => JXSSFSheet}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet} // superclass
    import java.lang.{Exception => JException}
    import java.io.{FileInputStream => JFileInputStream}
    import java.io.{File => JFile}
    import java.util.{Iterator => JIterator}

    pub enum XssfWorkbook(JXSSFWorkbook)

    pub eff Workbook {
        // `askXssfWorkbookHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askXssfWorkbookHandle(): XssfWorkbook
        def createSheet(): XssfSheet
        def createSheetWithName(s: String): XssfSheet
        def getNumberOfSheets(): Int32
        def getSheet(s: String): Option[XssfSheet]
        def getSheetAt(i: Int32): Option[XssfSheet]
        def getSheetIndex(s: String): Option[Int32]
        def getSheetIndexOfSheet(s: XssfSheet): Option[Int32]
        def getSheetName(i: Int32): String
        def isMacroEnabled(): Bool
        def setSheetName(i: Int32, s: String): Unit
    }


    pub def handle(wb: XssfWorkbook, 
                    f: a -> b \ ef): a -> Result[String, b] \ {IO + ef} - Workbook = x -> {
        let XssfWorkbook.XssfWorkbook(jworkbook) = wb;
        run {
            Ok(f(x))
        } with handler Workbook {
            
            def askXssfWorkbookHandle(k) = { k(wb) }

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
                case ex: JException => Err(ex.getMessage())
            }

            def getSheet(s, k) = try {
                let ans = jworkbook.getSheet(s);
                if (Object.isNull(ans)) 
                    k(None)
                else 
                    k(Some(XssfSheet.XssfSheet(ans)))                
            } catch {
                case ex: JException => Err(ex.getMessage())
            }


            def getSheetAt(i, k) = try {
                let ans = jworkbook.getSheetAt(i);
                k(Some(XssfSheet.XssfSheet(ans)))
            } catch {
                case _: JException => k(None)
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

            // TODO check failure mode...
            def getSheetName(i, k) = {
                let ans = jworkbook.getSheetName(i);
                k(ans)
            }

            def isMacroEnabled(k) = {
                    let ans = jworkbook.isMacroEnabled();
                    k(ans)
            }

            def setSheetName(i, s, k) = try {
                jworkbook.setSheetName(i, s);
                k()
            } catch {
                case ex: JException => Err(ex.getMessage())
            }
        }
    }

    pub def runReadOnly(path: String, f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - Workbook = try {
        let jfile = new JFile(path);
        let jfis = new JFileInputStream(jfile);        
        let jworkbook = new JXSSFWorkbook(jfis);
        let ans = handle(XssfWorkbook.XssfWorkbook(jworkbook), f)();
        jworkbook.close();
        jfis.close();
        ans
    } catch {
        case ex: JException => Err(ex.getMessage())
    }

    pub def runReadWrite(path: String, f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - Workbook = try {
        let jworkbook = new JXSSFWorkbook(path);
        let ans = handle(XssfWorkbook.XssfWorkbook(jworkbook), f)();
        jworkbook.close();
        ans
    } catch {
        case ex: JException => Err(ex.getMessage())
    }


    pub def getSheetNames(): List[String] \ Workbook = {
        let count = Workbook.getNumberOfSheets();
        List.unfold(i -> {
            if (i < count) {
                let name = Workbook.getSheetName(i);
                Some((name, i + 1))
            } else None
            }, 0)
    }

    pub def withSheet(name: String, f: Unit -> a \ ef): Result[String, a] \ (ef + Workbook) - Sheet = 
        match Workbook.getSheet(name) {
            case Some(s) => SheetIO.XSSF.Sheet.runSheet(s, f)
            case None    => Err("No Sheet")
        }

    pub def withSheetAt(ix: Int32, f: Unit -> a \ ef): Result[String, a] \ (ef + Workbook) - Sheet = 
        match Workbook.getSheetAt(ix) {
            case Some(s) => SheetIO.XSSF.Sheet.runSheet(s, f)
            case None    => Err("No Sheet")
        }

    pub def sheetIterator(rc: Region[r]): Iterator[XssfSheet, r, r] \ Workbook + r = 
        let XssfWorkbook.XssfWorkbook(jworkbook) = Workbook.askXssfWorkbookHandle();
        let javaIter: JIterator[JSheet] = unsafe IO as r { jworkbook.sheetIterator() };
        def wrap(jsheet) = { 
            let jxssfsheet = unchecked_cast(jsheet as JXSSFSheet);
            XssfSheet.XssfSheet(jxssfsheet)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JSheet]), javaIter)
            |> Iterator.map(wrap)

    pub def forEachSheet(f: XssfSheet -> Unit \ ef): Unit \ Workbook + ef = region rc {
        Iterator.forEach(f, sheetIterator(rc))
    }
    
    pub def withEachSheet(f: Unit -> Unit \ ef): Unit \ Workbook + ef - Sheet = region rc {
        Iterator.forEach(s -> discard SheetIO.XSSF.Sheet.runSheet(s, f),  sheetIterator(rc))
    }

}