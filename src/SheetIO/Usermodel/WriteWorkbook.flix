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

    use SheetIO.Usermodel.{Workbook, Sheet, CellStyle}
    use SheetIO.Usermodel.Name 

    import org.apache.poi.ss.usermodel.{Workbook => JWorkbook}
    import org.apache.poi.xssf.usermodel.{XSSFWorkbook => JXSSFWorkbook}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet} // superclass
    import org.apache.poi.openxml4j.util.{ZipSecureFile => JZipSecureFile}
    import java.lang.{Exception => JException}
    import java.io.{FileOutputStream => JFileOutputStream}
    import java.util.{Iterator => JIterator}


    pub eff WriteWorkbook {
        def createSheet(): Sheet
        def createSheetWithName(s: String): Sheet
        def cloneSheet(num: Int32): Sheet
        def removeSheetAt(i: Int32): Unit
        def createCellStyle(): CellStyle
        def createName(): Name
    }



    pub def handle(wb: Workbook,
                    f: a -> b \ ef): a -> Result[String, b] \ {IO + ef} - WriteWorkbook - Fail = x -> {
        let jworkbook = ToJava.toJava(wb);
        run {
            Ok(f(x))
        } with handler WriteWorkbook {

            def createSheet(k) = {
                let ans = jworkbook.createSheet();
                k(ToFlix.toFlix(ans))
            }
           
            def createSheetWithName(s, k) = {
                let ans = jworkbook.createSheet(s);
                k(ToFlix.toFlix(ans))
            }

            def cloneSheet(i, k) = {
                let ans = jworkbook.cloneSheet(i);
                k(ToFlix.toFlix(ans))
            }

            def removeSheetAt(i, k) = try {
                jworkbook.removeSheetAt(i);
                k()
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }

            def createCellStyle(k) = try {
                let jcs = jworkbook.createCellStyle();
                k(ToFlix.toFlix(jcs))
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }

            def createName(k) = try {
                let x1 = jworkbook.createName();
                k(ToFlix.toFlix(x1))
            } catch {
                case ex: JException => Fail.report(ex.getMessage())
            }

        } with handler Fail {
            def report(s, _) = Err(s)
        }
    }


    // pub def withSheet(name: String, f: Unit -> a \ ef): Result[String, a] \ (ef + WriteWorkbook) - WriteSheet - Alt =
    //     let s = ReadWorkbook.getSheet(name);
    //     SheetIO.Usermodel.ReadSheet.runReadSheetWithResult(s, f)


    // with EffectExtras.Fail.runWithResultOfIoError

    // pub def runReadWriteXssf(path: String, f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - WriteWorkbook - Fail = try {
    //     let jworkbook = new JXSSFWorkbook(path);
    //     let jworkbook1: JWorkbook = checked_cast(jworkbook);
    //     let ans = handle(ToFlix.toFlix(jworkbook1), f)();
    //     jworkbook.close();
    //     ans
    // } catch {
    //     case ex: JException => Err(ex.getMessage())
    // }

    pub def runWriteXssf(path: String, 
                        f: Unit -> a \ ef): Result[String, a] \ {IO + ef} - WriteWorkbook - Fail = try {
        let outs = new JFileOutputStream(path);
        let jworkbook = new JXSSFWorkbook();
        let jworkbook1: JWorkbook = checked_cast(jworkbook);
        let ans = handle(ToFlix.toFlix(jworkbook1), f)();
        jworkbook.write(outs);
        jworkbook.close();
        ans
    } catch {
        case ex: JException => Err(ex.getMessage())
    }


}
