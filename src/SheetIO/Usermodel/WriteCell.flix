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


pub mod SheetIO.Usermodel.WriteCell {
    
    
    use EffectExtras.Fail.Fail
    use InteropTime.LocalDate
    use InteropTime.LocalDateTime

    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.Comment.{Comment}
    use SheetIO.Usermodel.{Cell, CellStyle}

    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import java.lang.{Exception => JException}


    pub eff WriteCell {
        def removeCellComment(): Unit
        def setCellStyle(s: CellStyle): Unit
        def setComment(c: Comment): Unit
        def setBlank(): Unit
        def setCellValueBool(x: Bool): Unit
        def setCellValueFloat64(d: Float64): Unit
        def setCellValueString(s: String): Unit
        def setCellValueLocalDate(d: LocalDate): Unit
        def setCellValueLocalDateTime(d: LocalDateTime): Unit
    }

    pub def handle(cell: Cell, 
                    f: a -> b \ ef): a -> b \ (ef + Fail - WriteCell) = x -> 
        let jcell = ToJava.toJava(cell);
        run {
            f(x)
        } with handler WriteCell {

            def removeCellComment(k) = {
                unsafe IO as {ef + Fail - WriteCell} { jcell.removeCellComment() };
                k()
            }

            def setCellStyle(style, k) = {
                let style1 = ToJava.toJava(style);
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellStyle(style1) };
                k()
            }

            def setComment(comment, k) = {
                let comment1 = ToJava.toJava(comment);
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellComment(comment1) };
                k()
            }

            def setBlank(k) = {
                unsafe IO as {ef + Fail - WriteCell} { jcell.setBlank() };
                k()
            }

            def setCellValueBool(b, k) = {
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellValue(b) };
                k()
            }

            def setCellValueFloat64(d, k) = {
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellValue(d) };
                k()
            }

            def setCellValueString(s, k) = {
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellValue(s) };
                k()
            }

            def setCellValueLocalDate(d, k) = {
                let d1 = ToJava.toJava(d);
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellValue(d1) };
                k()
            }

            def setCellValueLocalDateTime(dt, k) = {
                let dt1 = ToJava.toJava(dt);
                unsafe IO as {ef + Fail - WriteCell} { jcell.setCellValue(dt1) };
                k()
            }
            
        }


    pub def runWriteCell(cell: Cell, 
                        f: Unit -> a \ ef): a \ (ef + Fail - WriteCell) = 
        handle(cell, f)()

    pub def runWriteCellWithResult(cell: Cell, 
                                    f: Unit -> a \ ef): Result[String, a] \ (ef - Fail - WriteCell) = 
        run { handle(cell, f)() } with EffectExtras.Fail.runWithResult

}
