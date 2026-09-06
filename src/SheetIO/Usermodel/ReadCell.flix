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


pub mod SheetIO.Usermodel.ReadCell {
    
    
    use EffectExtras.Alt.Alt
    use EffectExtras.Fail.Fail
    use InteropTime.LocalDateTime

    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellAddress.ToCellAddress
    use SheetIO.Common.CellReference
    use SheetIO.Common.CellReference.ToCellReference
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.{Cell, Comment}
    use SheetIO.Usermodel.ReadWorkbook
    use SheetIO.Usermodel.ReadSheet
    use SheetIO.Usermodel.ReadRow

    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import java.lang.{Exception => JException}


    pub eff ReadCell {
        def getAddress(): CellAddress
        def getCellType(): CellType
        def getCellValueBool(): Bool
        def getCellValueFloat64(): Float64
        def getCellValueString(): String
        def getCellValueLocalDateTime(): LocalDateTime
        def getCellFormula(): String
        def getCellComment(): Comment
        def getColumnIndex(): Int32
        def getRowIndex(): Int32
        def toString(): String
    }

    pub def handle(cell: Cell, 
                    f: a -> b \ ef): a -> b \ (ef + Fail - ReadCell) = x -> 
        let jcell = ToJava.toJava(cell);
        run {
            f(x)
        } with handler ReadCell {

            def getAddress(k) = {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getAddress() };
                EffectExtras.Fail.assertNotNull(ans);
                k(ToFlix.toFlix(ans))
            }

            def getCellType(k) = {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getCellType() };
                EffectExtras.Fail.assertNotNull(ans);
                k(ToFlix.toFlix(ans))
            }

            def getCellValueBool(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getBooleanCellValue() };
                k(ans)
            } catch {
                case _: JException => Fail.report("getBoolCellValue")
            }

            def getCellValueFloat64(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getNumericCellValue() };
                k(ans)
            } catch {
                case _: JException => Fail.report("getFloat64CellValue")
            }

            def getCellValueString(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getStringCellValue() };
                k(ans)
            } catch {
                case _: JException => Fail.report("getCellValueString")
            }

            def getCellValueLocalDateTime(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getLocalDateTimeCellValue() };
                k(ToFlix.toFlix(ans))
            } catch {
                case _: JException => Fail.report("getCellValueLocalDateTime")
            }


            def getCellFormula(k) = try {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getCellFormula() };
                k(ans)
            } catch {
                case _: JException => Fail.report("getCellFormula")
            }

            def getCellComment(k) = {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getCellComment() };
                EffectExtras.Fail.assertNotNull(ans);
                k(ToFlix.toFlix(ans))
            }

            def getColumnIndex(k) = {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getColumnIndex() };
                k(ans)
            } 

            def getRowIndex(k) = {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.getRowIndex() };
                k(ans)
            } 

            def toString(k) = {
                let ans = unsafe IO as {ef + Fail - ReadCell} { jcell.toString() };
                k(ans)
            } 

        }

    // The `ReadCell` effect uses the `Alt` rather than `Fail` effect because 
    // we expect a user may want to retry reading with the `Alt` combinators.
    // However when it is run failure is dispatched to the `Fail` effect so it
    // can be nested within other parts of the "Read" API which use Fail as the 
    // failure effect.
    pub def runReadCell(cell: Cell, f: Unit -> a \ ef): a \ (ef + Fail - ReadCell - Alt) = 
        run { handle(cell, f)() } with EffectExtras.Alt.runAlt





}
