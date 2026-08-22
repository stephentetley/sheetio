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


pub mod SheetIO.Usermodel.WriteCellStyle {
    
    
    use EffectExtras.Fail.Fail

    use SheetIO.Usermodel.{CellStyle, FillPatternType}

    import org.apache.poi.ss.usermodel.{CellStyle => JCellStyle}
    import java.lang.{Exception => JException}


    pub eff WriteCellStyle {
        def setFillBackgroundColor(s: Int16): Unit
        def setFillForegroundColor(s: Int16): Unit
        def setFillPattern(fpt: FillPatternType): Unit
    }

    pub def handle(cellstyle: CellStyle, 
                    f: a -> b \ ef): a -> b \ (ef + Fail - WriteCellStyle) = x -> 
        let jcellstyle = ToJava.toJava(cellstyle);
        run {
            f(x)
        } with handler WriteCellStyle {

            def setFillBackgroundColor(bg, k) = {
                unsafe IO as {ef + Fail - WriteCellStyle} { jcellstyle.setFillBackgroundColor(bg) };
                k()
            }

            def setFillForegroundColor(bg, k) = {
                unsafe IO as {ef + Fail - WriteCellStyle} { jcellstyle.setFillForegroundColor(bg) };
                k()
            }

            def setFillPattern(fpt, k) = {
                let fpt1 = ToJava.toJava(fpt);
                unsafe IO as {ef + Fail - WriteCellStyle} { jcellstyle.setFillPattern(fpt1) };
                k()
            }

            
        }


    pub def runWriteCellStyle(cellstyle: CellStyle, 
                        f: Unit -> a \ ef): a \ (ef + Fail - WriteCellStyle) = 
        handle(cellstyle, f)()

    pub def runWriteCellStyleWithResult(cellstyle: CellStyle, 
                                    f: Unit -> a \ ef): Result[String, a] \ (ef - Fail - WriteCellStyle) = 
        run { handle(cellstyle, f)() } with EffectExtras.Fail.runWithResult

}
