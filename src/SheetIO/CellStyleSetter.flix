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

/*
    pub enum CellStyleSetter[r: Region, ka: Type, a: Type](
        Graded.MutStateError[r, ka, Interop.Apache.Poi.SS.Usermodel.CellStyle[r], a]
    )


    instance GradedFunctor[CellStyleSetter[r, ka]] {
        pub def map(f: a -> b \ ef, o: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, b] \ ef = 
            SheetIO.CellStyleSetter.map(f, o)
    }

    instance GradedApplicative[CellStyleSetter[r, ka]] {
        pub def point(x: a): CellStyleSetter[r, ka, a] = 
            SheetIO.CellStyleSetter.point(x)

        pub def ap(f: CellStyleSetter[r, ka, a -> b \ ef], x: CellStyleSetter[r, ka, a]) : CellStyleSetter[r, ka, b] \ ef = 
            SheetIO.CellStyleSetter.ap(f, x)
    }

    instance GradedMonad[CellStyleSetter[r, ka]] {
        pub def flatMap(f: a -> CellStyleSetter[r, ka, b] \ ef, o: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, b] \ ef =
            SheetIO.CellStyleSetter.flatMap(f, o)
    }

    instance GradedError[CellStyleSetter[r, ka]] {
        pub def throwError(msg: String): CellStyleSetter[r, ka, a] = 
            SheetIO.CellStyleSetter.throwError(msg)

        pub def mapError(f: String -> String \ ef, ma: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, a] \ ef =
            SheetIO.CellStyleSetter.mapError(f, ma)

        pub def catchError(ma: CellStyleSetter[r, ka, a], f: String -> CellStyleSetter[r, ka, a] \ ef): CellStyleSetter[r, ka, a] \ ef =
            SheetIO.CellStyleSetter.catchError(ma, f)
    }
*/

}

mod SheetIO.CellStyleSetter {
/*
    use SheetIO.CellStyleSetter
    use SheetIO.CellStyleSetter.{CellStyleSetter}

    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Interop.Apache.Poi.SS.Usermodel.Font
    use Interop.Apache.Poi.SS.Usermodel.Font.{Font}


    pub def getCellStyleSetter(x: CellStyleSetter[r, ka, a]): MonadLib.MutStateError[r, ka, CellStyle[r], a] = 
        let CellStyleSetter(x1) = x;
        x1

    /// The state param `st` is expected to be mutated...
    pub def runCellStyleSetter!(r: Region[r], 
                                ma: CellStyleSetter[r, a, a], 
                                st: CellStyle[r]): Result[a, String] \ r =
        MonadLib.MutStateError.runMutStateError(r, getCellStyleSetter(ma), st)

    pub def point(x: a): CellStyleSetter[r, ka, a] = 
        CellStyleSetter(MonadLib.MutStateError.point(x))

    pub def map(f: a -> b \ ef, x: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, b] \ ef =
        CellStyleSetter(MonadLib.MutStateError.map(f, getCellStyleSetter(x)))

    pub def ap(mf: CellStyleSetter[r, ka, a -> b \ ef], ma: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, b] \ ef =
        CellStyleSetter(MonadLib.MutStateError.ap(getCellStyleSetter(mf), getCellStyleSetter(ma)))

    pub def flatMap(k: a -> CellStyleSetter[r, ka, b] \ ef, x: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, b]  \ ef = 
        CellStyleSetter(MonadLib.MutStateError.flatMap(k >> getCellStyleSetter, getCellStyleSetter(x)))


    pub def throwError(msg: String): CellStyleSetter[r, ka, a] = 
        CellStyleSetter(MonadLib.MutStateError.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: CellStyleSetter[r, ka, a]): CellStyleSetter[r, ka, a] \ ef =
        CellStyleSetter(MonadLib.MutStateError.mapError(f, getCellStyleSetter(ma)))

    pub def catchError(ma: CellStyleSetter[r, ka, a], f: String -> CellStyleSetter[r, ka, a] \ ef): CellStyleSetter[r, ka, a] \ ef =
        CellStyleSetter(MonadLib.MutStateError.catchError(getCellStyleSetter(ma), f >> getCellStyleSetter))

    def liftSetter!(f: CellStyle[r] -> a \ ef): CellStyleSetter[r, ka, a] \ ef  =
        CellStyleSetter(MonadLib.MutStateError.liftSetter!(f))

    // def liftSetterResult!(f: CellStyle[r] -> Result[a, err] \ ef): CellStyleSetter[r, ka, a] \ ef with ToString[err] =
    //     CellStyleSetter(MonadLib.MutStateError.liftSetterResult!(f))
    
    pub def liftAction(f: Unit -> a \ ef): CellStyleSetter[r, ka, a] \ ef =
        CellStyleSetter(MonadLib.MutStateError.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): CellStyleSetter[r, ka, a] \ ef with ToString[err] =
        CellStyleSetter(MonadLib.MutStateError.liftActionResult(f))


    pub def setFont!(font: Font[r]): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setFont!(font))


    pub def setWrapText!(wrapped: Bool): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setWrapText!(wrapped))


    pub def setIndention!(indent: Int16): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setIndention!(indent))

    pub def setFillForegroundColor!(color: Int16): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setFillForegroundColor!(color))
    
    pub def setFillBackgroundColor!(color: Int16): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setFillBackgroundColor!(color))


    pub def setShrinkToFit!(shrink: Bool): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setShrinkToFit!(shrink))

    pub def setDataFormat!(value: Int16): CellStyleSetter[r, ka, Unit] \ Write(r) =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setDataFormat!(value))


    // pub def cloneStyleFrom!(x: CellStyle[r]): CellStyleSetter[r, ka, Unit] \ Write(r) =
    //     let CellStyle(x1) = x;
    //     liftPrim(r -> SheetIO.Internal.CellStyle.cloneStyleFrom!(r, x1))

*/

}
