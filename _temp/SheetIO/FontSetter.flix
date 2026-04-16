/*
 * Copyright 2020 Stephen Tetley
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


mod SheetIO {

    pub type alias FontSetter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.Builder[ka, SheetIO.Internal.Apache.Poi.SS.Usermodel.Font[r], ef, a]

}

mod SheetIO.FontSetter {

    use Graded.Builder.{liftSetter!, liftSetterResult!}

    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Font
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.{Font}
    use SheetIO.FontSetter


    /// The state param `st` is expected to be mutated...
    pub def runFontSetter!(ma: FontSetter[a, r, ef, a], 
                            st: Font[r]): Result[String, a] \ ef =
        Graded.Builder.runBuilder(ma, st)



    pub def setFontName!(name: String): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setFontName!(name))

    pub def setFontHeight!(height: Int16): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setFontHeight!(height))

    pub def setFontHeightInPoints!(height: Int16): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setFontHeightInPoints!(height))

    pub def setBold!(bold: Bool): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setBold!(bold))

    pub def setItalic!(italic: Bool): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setItalic!(italic))

    pub def setStrikeout!(strikeout: Bool): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setStrikeout!(strikeout))
    
    pub def setColor!(color: Int16): FontSetter[ka, r, r, Unit] = 
        liftSetter!(SheetIO.Internal.Apache.Poi.SS.Usermodel.Font.setColor!(color))


    /// TODO others need enums...

}
