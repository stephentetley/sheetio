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


pub mod SheetIO.Usermodel.WriteFont {
    
    
    use EffectExtras.Alt.Alt
    use SheetIO.Usermodel.{Font}

    import org.apache.poi.ss.usermodel.{Font => JFont}
    import java.lang.{Exception => JException}


    pub eff WriteFont {
        def setBold(b: Bool): Unit
        def setItalic(b: Bool): Unit
        def setStrikeout(b: Bool): Unit
    }

    pub def handle(font: Font, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - WriteFont) = x -> 
        let jfont = ToJava.toJava(font);
        run {
            f(x)
        } with handler WriteFont {

            def setBold(b, k) = {
                unsafe IO as {ef + Alt - WriteFont} { jfont.setBold(b) };
                k()
            }

            def setItalic(b, k) = {
                unsafe IO as {ef + Alt - WriteFont} { jfont.setItalic(b) };
                k()
            }

            def setStrikeout(b, k) = {
                unsafe IO as {ef + Alt - WriteFont} { jfont.setStrikeout(b) };
                k()
            }
            
        }


    pub def runWriteFont(font: Font, 
                        f: Unit -> a \ ef): a \ (ef + Alt - WriteFont) = 
        handle(font, f)()

    pub def runWriteFontWithResult(font: Font, 
                                    f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - WriteFont) = 
        run { handle(font, f)() } with EffectExtras.Alt.runWithResult

}
