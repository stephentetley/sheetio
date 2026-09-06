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

    use SheetIO.Common.CellReference
    use SheetIO.Common.CellReference.ToCellReference
    use SheetIO.Usermodel.ReadSheet.ReadSheet
    use SheetIO.Usermodel.{Workbook, Sheet, Cell}
    use SheetIO.Usermodel.ReifyWorkbook

    import org.apache.poi.ss.usermodel.{Workbook => JWorkbook}
    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet} // superclass
    import org.apache.poi.openxml4j.util.{ZipSecureFile => JZipSecureFile}
    import java.lang.{Exception => JException}
    import java.io.{FileInputStream => JFileInputStream}
    import java.io.{File => JFile}
    import java.util.{Iterator => JIterator}


    pub eff ReadWorkbook {
        def getNumberOfSheets(): Int32
        def getSheet(s: String): Sheet
        def getSheetAt(i: Int32): Sheet
        def getSheetIndex(s: String): Option[Int32]
        def getSheetIndexOfSheet(s: Sheet): Option[Int32]
        def getSheetName(i: Int32): String
    }



    pub def handle(wb: Workbook,
                    f: a -> b \ ef): a -> b \ ef + IO + Fail - ReifyWorkbook - ReadWorkbook = x -> {
        let jworkbook = ToJava.toJava(wb);
        run {
            f(x)
        } with handler ReifyWorkbook {
            
            def getWorkbook(k) = k(wb)

        } with handler ReadWorkbook {
            
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
        }
    }

    // Change to:
    // ``f: Workbook -> a \ ef``
    // Doesn't have to use Readworkbook effect...
    pub def runReadWorkbookXssf(workbook: Workbook, 
                                f: Unit -> a \ ef): a \ ef + IO + Fail - ReifyWorkbook - ReadWorkbook =
        handle(workbook,f)()


    pub def withWorkbookXssf(path: String, 
                                min_inflate_ratio: Option[Float64],
                                f: Workbook -> a \ ef
                                ): a \ ef + IO + Fail= 
        try {
            let jfile = new JFile(path);
            let jfis = new JFileInputStream(jfile);
            match min_inflate_ratio {
                case Some(d) =>  JZipSecureFile.setMinInflateRatio(d) 
                case None    => ()
            };
            let jworkbook = new JXSSFWorkbook(jfis);
            let jworkbook1: JWorkbook = checked_cast(jworkbook);
            let ans = f(ToFlix.toFlix(jworkbook1));
            jworkbook.close();
            jfis.close();
            ans
        } catch {
            case ex: JException => Fail.report(ex.getMessage())
        }


    pub def getCellValueBool(cellRef: cellref): Bool \ ReadWorkbook + Fail with ToCellReference[cellref] = 
        let ref1: CellReference = ToCellReference.toCellReference(cellRef) |> Fail.fromOption;
        let sname = CellReference.getSheetName(ref1);
        let sheet = ReadWorkbook.getSheet(sname);
        SheetIO.Usermodel.ReadSheet.runReadSheet(sheet, _ -> {
            SheetIO.Usermodel.ReadSheet.getCellValueBool(CellReference.toCellAddress(ref1))
        })

    pub def getCellValueFloat64(cellRef: cellref): Float64 \ ReadWorkbook + Fail with ToCellReference[cellref] = 
        let ref1: CellReference = ToCellReference.toCellReference(cellRef) |> Fail.fromOption;
        let sname = CellReference.getSheetName(ref1);
        let sheet = ReadWorkbook.getSheet(sname);
        SheetIO.Usermodel.ReadSheet.runReadSheet(sheet, _ -> {
            SheetIO.Usermodel.ReadSheet.getCellValueFloat64(CellReference.toCellAddress(ref1))
        })

    pub def getCellValueString(cellRef: cellref): String \ ReadWorkbook + Fail with ToCellReference[cellref] = 
        let ref1: CellReference = ToCellReference.toCellReference(cellRef) |> Fail.fromOption;
        let sname = CellReference.getSheetName(ref1);
        let sheet = ReadWorkbook.getSheet(sname);
        SheetIO.Usermodel.ReadSheet.runReadSheet(sheet, _ -> {
            SheetIO.Usermodel.ReadSheet.getCellValueString(CellReference.toCellAddress(ref1))
        })



    pub def getSheetNames(): List[String] \ ReadWorkbook = {
        let count = ReadWorkbook.getNumberOfSheets();
        List.unfold(i -> {
            if (i < count) {
                let name = ReadWorkbook.getSheetName(i);
                Some((name, i + 1))
            } else None
            }, 0)
    }

    pub def sheetIterator(rc: Region[r], workbook: Workbook): Iterator[Sheet, r, r] \ r =
        let jworkbook = ToJava.toJava(workbook);
        let javaIter: JIterator[JSheet] = unsafe IO as r { jworkbook.sheetIterator() };
        def wrap(jsheet) = { 
            ToFlix.toFlix(jsheet)
        };
        Adaptor.fromIterator(rc, (Proxy.Proxy: Proxy[JSheet]), javaIter)
            |> Iterator.map(wrap)
    
    pub def forEachSheet(f: Sheet -> Unit \ ef, t: Workbook): Unit \ ef = region rc {
        Iterator.forEach(f, sheetIterator(rc, t))
    }
    

    pub def foldSheets(f: Sheet -> st -> st \ ef, state: st): st \ ef + ReifyWorkbook + Fail = region rc {
        let workbook = ReifyWorkbook.getWorkbook();
        let iter = sheetIterator(rc, workbook);
        @Tailrec
        def loop(st) = match Iterator.next(iter) {
            case None       => st
            case Some(sheet) => match EffectExtras.Fail.runWithResult(_ -> f(sheet, st)) {
                case Err(msg)   => Fail.report(msg)
                case Ok(st1)    => loop(st1)
            }
        };
        loop(state)
    }

    pub def filterMapSheets(f: Sheet -> Option[a] \ ef): List[a] \ ef + ReifyWorkbook + Fail = region rc {
        let workbook = ReifyWorkbook.getWorkbook();
        let iter = sheetIterator(rc, workbook);
        @Tailrec
        def loop(k) = match Iterator.next(iter) {
            case None       => k(Nil)
            case Some(sheet) => match EffectExtras.Fail.runWithResult(_ -> f(sheet)) {
                case Err(msg)       => Fail.report(msg)
                case Ok(None)       => loop(k)
                case Ok(Some(ans))  => loop(ks -> k(ans :: ks))
            }
        };
        loop(xs -> xs)
    }


    pub def getCellByReference(cr: CellReference): Cell \ ReadWorkbook + Fail = 
        let sheetname = CellReference.getSheetName(cr);
        let sheet = ReadWorkbook.getSheet(sheetname);
        let celladdr = CellReference.toCellAddress(cr);
        SheetIO.Usermodel.ReadSheet.runReadSheet(sheet, _ -> {
            ReadSheet.getCell(celladdr)
        })


}

