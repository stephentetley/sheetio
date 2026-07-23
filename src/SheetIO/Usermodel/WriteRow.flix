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


pub mod SheetIO.Usermodel.WriteRow {

    use EffectExtras.Alt.Alt


    use SheetIO.Usermodel.{Sheet, Row, Cell}

    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell} // superclass
    import java.lang.{Exception => JException}
    import java.util.{Iterator => JIterator}

    pub eff WriteRow {
        // `askRowHandle` should not be used in user code
        // It is a hack until we can run local effects polymorphically
        def askRowHandle(): Row
        def createCell(i: Int32): Cell
    }

        
    pub def handle(row: Row, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - WriteRow) + r  = x -> {
        let jrow = ToJava.toJava(row);
        run {
            f(x)
        } with handler WriteRow {
            
            def askRowHandle(k) = k(row)

            def createCell(i, k) = {
                let ans = unsafe IO as {ef + Alt - WriteRow} { jrow.createCell(i) };
                if (not Object.isNull(ans)) 
                    k(ToFlix.toFlix(ans))
                else 
                    Alt.failure("createCell ${i}")
            }

        }
    }

    pub def runWriteRow(row: Row, f: Unit -> a \ ef): a \ (ef + Alt - WriteRow) = 
        handle(row, f)()

    pub def runWriteRowWithResult(row: Row, 
                                    f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - WriteRow) = 
        run { handle(row, f)() } with EffectExtras.Alt.runWithResult



}

