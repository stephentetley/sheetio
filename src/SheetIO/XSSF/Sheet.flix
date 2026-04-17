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


pub mod SheetIO.XSSF.Sheet {

    use IoError.ErrorKind
    use IoError.IoError

    import org.apache.poi.xssf.usermodel.{XSSFSheet => JXSSFSheet}
    import java.lang.{Exception => JException}

    pub enum XssfSheet(JXSSFSheet)

    pub eff Sheet {
        def getSheetName() : String
    }

    pub def handle(sheet: XssfSheet, 
                f: a -> b \ ef): a -> Result[String, b] \ (ef - Sheet) = x -> region rc {
        handleHelper(rc, sheet, f, x)
    }


    def handleHelper(_: Region[r], 
                        sheet: XssfSheet, 
                        f: a -> b \ ef): a -> Result[String, b] \ (ef - Sheet) + r = x -> {
        let XssfSheet.XssfSheet(jsheet) = sheet;
        run {
            Ok(f(x))
        } with handler Sheet {
            
            def getSheetName(k) = try {
                let ans = unchecked_cast(jsheet.getSheetName() as _ \ r);
                k(ans)
            } catch {
                case ex: JException => Err(ex.getMessage())
            }
        }
    }

    pub def runSheet(sheet: XssfSheet, f: Unit -> a \ ef): Result[String, a] \ (ef - Sheet) = 
        handle(sheet, f)()

}