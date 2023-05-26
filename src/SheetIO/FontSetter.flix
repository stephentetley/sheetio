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


    pub enum FontSetter[r: Region, ka: Type, ef: Eff, a: Type](
        Graded.MutStateError[ka, Interop.Apache.Poi.SS.Usermodel.Font[r], ef, a]
    )


    instance GradedFunctor[FontSetter[r, ka]] {
        pub def map(f: a -> b \ ef1, 
                    o: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1, b] = 
            SheetIO.FontSetter.map(f, o)
    }

    instance GradedApplicative[FontSetter[r, ka]] {
        pub def point(x: a): FontSetter[r, ka, ef, a] = 
            SheetIO.FontSetter.point(x)

        pub def ap(f: FontSetter[r, ka, ef1, a -> b \ ef2], 
                    x: FontSetter[r, ka, ef, a]) : FontSetter[r, ka, ef + ef1 + ef2, b] = 
            SheetIO.FontSetter.ap(f, x)
    }

    instance GradedMonad[FontSetter[r, ka]] {
        pub def flatMap(f: a -> FontSetter[r, ka, ef1, b] \ ef2, 
                        o: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1 + ef2, b] =
            SheetIO.FontSetter.flatMap(f, o)
    }


    instance GradedError[FontSetter[r, ka]] {
        pub def throwError(msg: String): FontSetter[r, ka, ef, a] = 
            SheetIO.FontSetter.throwError(msg)

        pub def mapError(f: String -> String \ ef1, 
                            ma: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1, a] =
            SheetIO.FontSetter.mapError(f, ma)

        pub def catchError(ma: FontSetter[r, ka, ef, a], 
                            f: String -> FontSetter[r, ka, ef1, a] \ ef2): FontSetter[r, ka, ef + ef1 + ef2, a] =
            SheetIO.FontSetter.catchError(ma, f)

        pub def reifyError(ma: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef, Result[String, a]] =
            SheetIO.FontSetter.reifyError(ma)
    }

}

mod SheetIO.FontSetter {

    use SheetIO.FontSetter
    use SheetIO.FontSetter.FontSetter
    use Interop.Apache.Poi.SS.Usermodel.Font
    use Interop.Apache.Poi.SS.Usermodel.Font.{Font}


    pub def getFontSetter(x: FontSetter[r, ka, ef, a]): Graded.MutStateError[ka, Font[r], ef, a] = 
        let FontSetter(x1) = x;
        x1

    /// The state param `st` is expected to be mutated...
    pub def runFontSetter!(ma: FontSetter[r, a, ef, a], 
                            st: Font[r]): Result[String, a] \ ef =
        Graded.MutStateError.runMutStateError(getFontSetter(ma), st)

    pub def point(x: a): FontSetter[r, ka, ef, a] = 
        FontSetter(Graded.MutStateError.point(x))

    pub def map(f: a -> b \ ef1, x: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1, b]  =
        FontSetter(Graded.MutStateError.map(f, getFontSetter(x)))

    pub def ap(mf: FontSetter[r, ka, ef1, a -> b \ ef2], 
                ma: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1 + ef2, b] =
        FontSetter(Graded.MutStateError.ap(getFontSetter(mf), getFontSetter(ma)))

    pub def flatMap(k: a -> FontSetter[r, ka, ef1, b] \ ef2, 
                    x: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1 + ef2, b] = 
        FontSetter(Graded.MutStateError.flatMap(k >> getFontSetter, getFontSetter(x)))


    pub def throwError(msg: String): FontSetter[r, ka, ef, a] = 
        FontSetter(Graded.MutStateError.throwError(msg))

    pub def mapError(f: String -> String \ ef1, 
                        ma: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef + ef1, a] =
        FontSetter(Graded.MutStateError.mapError(f, getFontSetter(ma)))

    pub def catchError(ma: FontSetter[r, ka, ef, a], 
                        f: String -> FontSetter[r, ka, ef1, a] \ ef2): FontSetter[r, ka, ef + ef1 + ef2, a] =
        FontSetter(Graded.MutStateError.catchError(getFontSetter(ma), f >> getFontSetter))

    pub def reifyError(ma: FontSetter[r, ka, ef, a]): FontSetter[r, ka, ef, Result[String, a]] =
        FontSetter(Graded.MutStateError.reifyError(getFontSetter(ma)))

    def liftSetter!(f: Font[r] -> a \ ef): FontSetter[r, ka, ef, a]  =
        FontSetter(Graded.MutStateError.liftSetter!(f))

    // def liftSetterResult!(f: Font[r] -> Result[a, err] \ ef): FontSetter[r, ka, ef,a] \ ef with ToString[err] =
    //     FontSetter(Graded.MutStateError.liftSetterResult!(f))
/*    
    pub def liftAction(f: Unit -> a \ ef): FontSetter[r, ka, ef,a] \ ef =
        FontSetter(Graded.MutStateError.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): FontSetter[r, ka, ef,a] \ ef with ToString[err] =
        FontSetter(Graded.MutStateError.liftActionResult(f))
*/

    pub def setFontName!(name: String): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setFontName!(name))

    pub def setFontHeight!(height: Int16): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setFontHeight!(height))

    pub def setFontHeightInPoints!(height: Int16): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setFontHeightInPoints!(height))

    pub def setBold!(bold: Bool): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setBold!(bold))

    pub def setItalic!(italic: Bool): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setItalic!(italic))

    pub def setStrikeout!(strikeout: Bool): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setStrikeout!(strikeout))
    
    pub def setColor!(color: Int16): FontSetter[r, ka, r, Unit] = 
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.Font.setColor!(color))


    /// TODO others need enums...

}
