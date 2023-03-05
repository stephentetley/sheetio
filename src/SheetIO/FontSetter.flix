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


namespace SheetIO {


    pub enum FontSetter[r: Region, ka: Type, a: Type](
        Graded.MutStateError[r, ka, Interop/Apache/Poi/SS/Usermodel.Font[r], a]
    )

/*
    instance Functor[FontSetter[r, ka]] {
        pub def map(f: a -> b & ef, o: FontSetter[r, ka, a]): FontSetter[r, ka, b] & ef = 
            SheetIO/FontSetter.map(f, o)
    }

    instance Applicative[FontSetter[r, ka]] {
        pub def point(x: a): FontSetter[r, ka, a] = 
            SheetIO/FontSetter.point(x)

        pub def ap(f: FontSetter[r, ka, a -> b & ef], x: FontSetter[r, ka, a]) : FontSetter[r, ka, b] & ef = 
            SheetIO/FontSetter.ap(f, x)
    }

    instance Monad[FontSetter[r, ka]] {
        pub def flatMap(f: a -> FontSetter[r, ka, b] & ef, o: FontSetter[r, ka, a]): FontSetter[r, ka, b] & ef =
            SheetIO/FontSetter.flatMap(f, o)
    }


    instance MonadLib.MonadError[FontSetter[r, ka]] {
        pub def throwError(msg: String): FontSetter[r, ka, a] = 
            SheetIO/FontSetter.throwError(msg)

        pub def mapError(f: String -> String \ ef, ma: FontSetter[r, ka, a]): FontSetter[r, ka, a] \ ef =
            SheetIO/FontSetter.mapError(f, ma)

        pub def catchError(ma: FontSetter[r, ka, a], f: String -> FontSetter[r, ka, a] \ ef): FontSetter[r, ka, a] \ ef =
            SheetIO/FontSetter.catchError(ma, f)
    }
*/

}

namespace SheetIO/FontSetter {

    use SheetIO.FontSetter
    use SheetIO.FontSetter.FontSetter
    use Interop/Apache/Poi/SS/Usermodel.Font
    use Interop/Apache/Poi/SS/Usermodel.Font.{Font}

/*
    pub def getFontSetter(x: FontSetter[r, ka, a]): MonadLib.MutStateError[r, ka, Font[r], a] = 
        let FontSetter(x1) = x;
        x1

    /// The state param `st` is expected to be mutated...
    pub def runFontSetter!(r: Region[r], 
                                ma: FontSetter[r, a, a], 
                                st: Font[r]): Result[a, String] \ r =
        MonadLib/MutStateError.runMutStateError(r, getFontSetter(ma), st)

    pub def point(x: a): FontSetter[r, ka, a] = 
        FontSetter(MonadLib/MutStateError.point(x))

    pub def map(f: a -> b \ ef, x: FontSetter[r, ka, a]): FontSetter[r, ka, b] \ ef =
        FontSetter(MonadLib/MutStateError.map(f, getFontSetter(x)))

    pub def ap(mf: FontSetter[r, ka, a -> b \ ef], ma: FontSetter[r, ka, a]): FontSetter[r, ka, b] \ ef =
        FontSetter(MonadLib/MutStateError.ap(getFontSetter(mf), getFontSetter(ma)))

    pub def flatMap(k: a -> FontSetter[r, ka, b] \ ef, x: FontSetter[r, ka, a]): FontSetter[r, ka, b]  \ ef = 
        FontSetter(MonadLib/MutStateError.flatMap(k >> getFontSetter, getFontSetter(x)))


    pub def throwError(msg: String): FontSetter[r, ka, a] = 
        FontSetter(MonadLib/MutStateError.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: FontSetter[r, ka, a]): FontSetter[r, ka, a] \ ef =
        FontSetter(MonadLib/MutStateError.mapError(f, getFontSetter(ma)))

    pub def catchError(ma: FontSetter[r, ka, a], f: String -> FontSetter[r, ka, a] \ ef): FontSetter[r, ka, a] \ ef =
        FontSetter(MonadLib/MutStateError.catchError(getFontSetter(ma), f >> getFontSetter))

    def liftSetter!(f: Font[r] -> a \ ef): FontSetter[r, ka, a] \ ef  =
        FontSetter(MonadLib/MutStateError.liftSetter!(f))

    // def liftSetterResult!(f: Font[r] -> Result[a, err] \ ef): FontSetter[r, ka, a] \ ef with ToString[err] =
    //     FontSetter(MonadLib/MutStateError.liftSetterResult!(f))
    
    pub def liftAction(f: Unit -> a \ ef): FontSetter[r, ka, a] \ ef =
        FontSetter(MonadLib/MutStateError.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): FontSetter[r, ka, a] \ ef with ToString[err] =
        FontSetter(MonadLib/MutStateError.liftActionResult(f))

    pub def setFontName!(name: String): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setFontName!(name))

    pub def setFontHeight!(height: Int16): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setFontHeight!(height))

    pub def setFontHeightInPoints!(height: Int16): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setFontHeightInPoints!(height))

    pub def setBold!(bold: Bool): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setBold!(bold))

    pub def setItalic!(italic: Bool): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setItalic!(italic))

    pub def setStrikeout!(strikeout: Bool): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setStrikeout!(strikeout))
    
    pub def setColor!(color: Int16): FontSetter[r, ka, Unit] \ Write(r) = 
        liftSetter!(Interop/Apache/Poi/SS/Usermodel/Font.setColor!(color))


    /// TODO others need enums...
*/
}
