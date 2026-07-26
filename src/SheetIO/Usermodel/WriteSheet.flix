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


pub mod SheetIO.Usermodel.WriteSheet {

    
    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail

    use SheetIO.Common.CellAddress.CellAddress
    use SheetIO.Usermodel.ReadRow.ReadRow
    use SheetIO.Usermodel.ReadCell.ReadCell
    use SheetIO.Usermodel.{Workbook, Sheet, Row, Cell}

    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import org.apache.poi.ss.util.{SheetUtil => JSheetUtil}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}


    pub eff WriteSheet {
        // `askSheetHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askSheetHandle(): Sheet
        def createRow(i: Int32) : Row
    }


    
    pub def handle(sheet: Sheet,
                    f: a -> b \ ef): a -> b \ (ef + Alt - WriteSheet) + r = x -> {
        let jsheet = ToJava.toJava(sheet);
        run {
            f(x)
        } with handler WriteSheet {
            
            def askSheetHandle(k) = { k(sheet) }

            def createRow(i, k) = {
                let ans = unsafe IO as {ef + Alt - WriteSheet} { jsheet.createRow(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("getRow ${i}")
            } 

        }
    }

    pub def runWriteSheet(sheet: Sheet, f: Unit -> a \ ef): a \ (ef + Alt - WriteSheet) =
        handle(sheet, f)()
    
    pub def runWriteSheetFail(sheet: Sheet, f: Unit -> a \ ef): a \ (ef + Fail - WriteSheet - Alt) =
        run { handle(sheet, f)() } with EffectExtras.Alt.runWithFail

    pub def runWriteSheetWithResult(sheet: Sheet, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - WriteSheet) =
        run { handle(sheet, f)() } with EffectExtras.Alt.runWithResult
    


}

