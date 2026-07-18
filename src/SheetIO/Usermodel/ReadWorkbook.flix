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


pub mod SheetIO.Usermodel.ReadWorkbook {
    
    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail

    use SheetIO.Usermodel.ReadSheet.ReadSheet
    use SheetIO.Usermodel.{Workbook, Sheet}

    import org.apache.poi.ss.usermodel.{Workbook => JWorkbook}
    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet} // superclass
    import org.apache.poi.openxml4j.util.{ZipSecureFile => JZipSecureFile}
    import java.lang.{Exception => JException}
    import java.io.{FileInputStream => JFileInputStream}
    import java.io.{File => JFile}
    import java.util.{Iterator => JIterator}


    pub eff ReadWorkbook {
        // `askWorkbookHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askWorkbookHandle(): Workbook
        def createSheet(): Sheet
        def createSheetWithName(s: String): Sheet
        def getNumberOfSheets(): Int32
        def getSheet(s: String): Sheet
        def getSheetAt(i: Int32): Sheet
        def getSheetIndex(s: String): Option[Int32]
        def getSheetIndexOfSheet(s: Sheet): Option[Int32]
        def getSheetName(i: Int32): String
    }



    pub def handle(wb: Workbook,
                    f: a -> b \ ef): a -> Result[String, b] \ {IO + ef} - ReadWorkbook - Fail = x -> {
        let jworkbook = ToJava.toJava(wb);
        run {
            Ok(f(x))
        } with handler ReadWorkbook {
            
            def askWorkbookHandle(k) = { k(wb) }

            def createSheet(k) = {
                let ans = jworkbook.createSheet();
                k(ToFlix.toFlix(ans))
            }
           
            def createSheetWithName(s, k) = {
                let ans = jworkbook.createSheet(s);
                k(ToFlix.toFlix(ans))
            }

            def getNumberOfSheets(k) = try {
                let ans = jworkbook.getNumberOfSheets();
                k(ans)
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }

            def getSheet(s, k) = try {
                let ans = jworkbook.getSheet(s);
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Fail.report("getSheet ${s} - sheet is null")
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }


            def getSheetAt(i, k) = try {
                let ans = jworkbook.getSheetAt(i);
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Fail.report("getSheetAt ${i} - sheet is null")
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }

            def getSheetIndex(s, k) = {
                let ans = jworkbook.getSheetIndex(s);
                if (ans < 0) k(None) else k(Some(ans))
            }

            def getSheetIndexOfSheet(s, k) = {
                let jsheet = ToJava.toJava(s);
                let ans = jworkbook.getSheetIndex(jsheet);
                if (ans < 0) k(None) else k(Some(ans))
            }

            // TODO check failure mode...
            def getSheetName(i, k) = {
                let ans = jworkbook.getSheetName(i);
                k(ans)
            }

        } with handler Fail {
            def report(s, _) = Err(s)
        }
    }

    pub def runReadWorkbookXssf(path: String, min_inflate_ratio: Option[Float64],  f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - ReadWorkbook - Fail = try {
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

    pub def runReadWriteXssf(path: String, f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - ReadWorkbook - Fail = try {
        let jworkbook = new JXSSFWorkbook(path);
        let jworkbook1: JWorkbook = checked_cast(jworkbook);
        let ans = handle(ToFlix.toFlix(jworkbook1), f)();
        jworkbook.close();
        ans
    } catch {
        case ex: JException => Err(ex.getMessage())
    }


    pub def getSheetNames(): List[String] \ ReadWorkbook = {
        let count = ReadWorkbook.getNumberOfSheets();
        List.unfold(i -> {
            if (i < count) {
                let name = ReadWorkbook.getSheetName(i);
                Some((name, i + 1))
            } else None
            }, 0)
    }

    pub def withSheet(name: String, f: Unit -> a \ ef): Result[String, a] \ (ef + ReadWorkbook) - ReadSheet - Alt =
        let s = ReadWorkbook.getSheet(name);
        SheetIO.Usermodel.ReadSheet.runReadSheetWithResult(s, f)

    pub def withSheetAt(ix: Int32, f: Unit -> a \ ef): Result[String, a] \ (ef + ReadWorkbook) - ReadSheet - Alt =
        let s = ReadWorkbook.getSheetAt(ix);
        SheetIO.Usermodel.ReadSheet.runReadSheetWithResult(s, f)

    pub def sheetIterator(rc: Region[r]): Iterator[Sheet, r, r] \ ReadWorkbook + r =
        let jworkbook = ReadWorkbook.askWorkbookHandle() |> ToJava.toJava;
        let javaIter: JIterator[JSheet] = unsafe IO as r { jworkbook.sheetIterator() };
        def wrap(jsheet) = { 
            ToFlix.toFlix(jsheet)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JSheet]), javaIter)
            |> Iterator.map(wrap)
    

    /// TODO ignores failure...
    pub def withEachSheet(f: Unit -> Unit \ ef): Unit \ ReadWorkbook + ef - ReadSheet - Alt = region rc {
        Iterator.forEach(s -> discard SheetIO.Usermodel.ReadSheet.runReadSheetWithResult(s, f),  sheetIterator(rc))
    }


    pub def foldSheets(f: st -> st \ ef, state: st): Result[String, st] \ ReadWorkbook + ef - ReadSheet - Alt = region rc {
        let iter = sheetIterator(rc);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None => Ok(st)
            case Some(sheet) => match SheetIO.Usermodel.ReadSheet.runReadSheetWithResult(sheet, _ -> f(st)) {
                case Err(msg) => Err(msg)
                case Ok(st1)  => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapSheets(f: Unit -> Option[a] \ ef): Result[String, List[a]] \ ReadWorkbook + ef - ReadSheet - Alt = region rc {
        let iter = sheetIterator(rc);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None => Ok(k(Nil))
            case Some(sheet) => match SheetIO.Usermodel.ReadSheet.runReadSheetWithResult(sheet, _ -> f()) {
                case Err(msg) => Err(msg)
                case Ok(None)  => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }

}
