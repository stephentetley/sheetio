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


pub mod SheetIO.Usermodel.WriteWorkbook {
    
    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail

    use SheetIO.Usermodel.{Workbook, Sheet}

    import org.apache.poi.ss.usermodel.{Workbook => JWorkbook}
    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet} // superclass
    import org.apache.poi.openxml4j.util.{ZipSecureFile => JZipSecureFile}
    import java.lang.{Exception => JException}
    import java.io.{FileInputStream => JFileInputStream}
    import java.io.{File => JFile}
    import java.util.{Iterator => JIterator}


    pub eff WriteWorkbook {
        // `askWorkbookHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askWorkbookHandle(): Workbook
        def createSheet(): Sheet
        def createSheetWithName(s: String): Sheet

    }



    pub def handle(wb: Workbook,
                    f: a -> b \ ef): a -> Result[String, b] \ {IO + ef} - WriteWorkbook - Fail = x -> {
        let jworkbook = ToJava.toJava(wb);
        run {
            Ok(f(x))
        } with handler WriteWorkbook {
            
            def askWorkbookHandle(k) = { k(wb) }

            def createSheet(k) = {
                let ans = jworkbook.createSheet();
                k(ToFlix.toFlix(ans))
            }
           
            def createSheetWithName(s, k) = {
                let ans = jworkbook.createSheet(s);
                k(ToFlix.toFlix(ans))
            }

        } with handler Fail {
            def report(s, _) = Err(s)
        }
    }

    pub def runReadWorkbookXssf(path: String, min_inflate_ratio: Option[Float64],  f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - WriteWorkbook - Fail = try {
        let jfile = new JFile(path);
        let jfis = new JFileInputStream(jfile);
        match min_inflate_ratio {
            case Some(d) =>  JZipSecureFile.setMinInflateRatio(d) 
            case None    => ()
        };
        let jworkbook = new JXSSFWorkbook(jfis);
        let jworkbook1: JWorkbook = checked_cast(jworkbook);
        let ans = handle(ToFlix.toFlix(jworkbook1), f)();
        jworkbook.close();
        jfis.close();
        ans
    } catch {
        case ex: JException => Err(ex.getMessage())
    }


    // with EffectExtras.Fail.runWithResultOfIoError

    pub def runReadWriteXssf(path: String, f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - WriteWorkbook - Fail = try {
        let jworkbook = new JXSSFWorkbook(path);
        let jworkbook1: JWorkbook = checked_cast(jworkbook);
        let ans = handle(ToFlix.toFlix(jworkbook1), f)();
        jworkbook.close();
        ans
    } catch {
        case ex: JException => Err(ex.getMessage())
    }



}
