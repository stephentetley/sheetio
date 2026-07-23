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
    
    
    use EffectExtras.Alt.Alt
    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.{Cell, Comment}

    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import java.lang.{Exception => JException}


    pub eff WriteCell {
        def setBlank(): Unit
    }

    pub def handle(cell: Cell, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - WriteCell) = x -> 
        let jcell = ToJava.toJava(cell);
        run {
            f(x)
        } with handler WriteCell {

            def setBlank(k) = {
                unsafe IO as {ef + Alt - WriteCell} { jcell.setBlank() };
                k()
            }

        }


    pub def runWriteCell(cell: Cell, 
                        f: Unit -> a \ ef): a \ (ef + Alt - WriteCell) = 
        handle(cell, f)()

    pub def runWriteCellWithResult(cell: Cell, 
                                    f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - WriteCell) = 
        run { handle(cell, f)() } with EffectExtras.Alt.runWithResult

}
