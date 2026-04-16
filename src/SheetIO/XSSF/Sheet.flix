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

    pub enum SheetHandle(JXSSFSheet)

    pub eff Sheet {
        def getSheetName() : String
    }

    pub def handle(sheet: SheetHandle, 
                f: a -> b \ ef): a -> Result[IoError, b] \ {IO + ef} - Sheet = x -> {
        let SheetHandle.SheetHandle(jsheet) = sheet;
        run {
            Ok(f(x))
        } with handler Sheet {
            
            def getSheetName(k) = try {
                let ans = jsheet.getSheetName();
                k(ans)
            } catch {
                case ex: JException => Err(IoError(ErrorKind.Other, ex.getMessage()))
            }
        }
    }

    pub def runWithIO(sheet: SheetHandle, f: Unit -> a \ ef): Result[IoError, a] \ {IO + ef} - Sheet = 
        handle(sheet, f)()

}