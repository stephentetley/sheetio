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


    pub enum CellStyleSetter[r: Region, ka: Type, ef: Eff, a: Type](
        Graded.MutStateError[ka, Interop.Apache.Poi.SS.Usermodel.CellStyle[r], ef, a]
    )


    instance GradedFunctor[CellStyleSetter[r, ka]] {
        pub def map(f: a -> b \ ef1, 
                    o: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1, b] = 
            SheetIO.CellStyleSetter.map(f, o)
    }

    instance GradedApplicative[CellStyleSetter[r, ka]] {
        pub def point(x: a): CellStyleSetter[r, ka, ef, a] = 
            SheetIO.CellStyleSetter.point(x)

        pub def ap(f: CellStyleSetter[r, ka, ef1, a -> b \ ef2], 
                    x: CellStyleSetter[r, ka, ef, a]) : CellStyleSetter[r, ka, ef + ef1 + ef2, b] = 
            SheetIO.CellStyleSetter.ap(f, x)
    }

    instance GradedMonad[CellStyleSetter[r, ka]] {
        pub def flatMap(f: a -> CellStyleSetter[r, ka, ef1, b] \ ef2, 
                        o: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1 + ef2, b] =
            SheetIO.CellStyleSetter.flatMap(f, o)
    }

    instance GradedError[CellStyleSetter[r, ka]] {
        pub def throwError(msg: String): CellStyleSetter[r, ka, ef, a] = 
            SheetIO.CellStyleSetter.throwError(msg)

        pub def mapError(f: String -> String \ ef1, 
                            ma: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1, a] =
            SheetIO.CellStyleSetter.mapError(f, ma)

        pub def catchError(ma: CellStyleSetter[r, ka, ef, a], 
                            f: String -> CellStyleSetter[r, ka, ef1, a] \ ef2): CellStyleSetter[r, ka, ef + ef1 + ef2, a] =
            SheetIO.CellStyleSetter.catchError(ma, f)

        pub def reifyError(ma: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef, Result[String, a]] =
            SheetIO.CellStyleSetter.reifyError(ma)
    }


}

mod SheetIO.CellStyleSetter {

    use SheetIO.CellStyleSetter
    use SheetIO.CellStyleSetter.{CellStyleSetter}

    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Interop.Apache.Poi.SS.Usermodel.Font
    use Interop.Apache.Poi.SS.Usermodel.Font.{Font}


    pub def getCellStyleSetter(x: CellStyleSetter[r, ka, ef, a]): Graded.MutStateError[ka, CellStyle[r], ef, a] = 
        let CellStyleSetter(x1) = x;
        x1

    /// The state param `st` is expected to be mutated...
    pub def runCellStyleSetter!(ma: CellStyleSetter[r, a, ef, a], 
                                st: CellStyle[r]): Result[String, a] \ ef =
        Graded.MutStateError.runMutStateError(getCellStyleSetter(ma), st)

    pub def point(x: a): CellStyleSetter[r, ka, ef, a] = 
        CellStyleSetter(Graded.MutStateError.point(x))

    pub def map(f: a -> b \ ef1, x: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1, b] =
        CellStyleSetter(Graded.MutStateError.map(f, getCellStyleSetter(x)))

    pub def ap(mf: CellStyleSetter[r, ka, ef1, a -> b \ ef2], 
                ma: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1 + ef2, b] =
        CellStyleSetter(Graded.MutStateError.ap(getCellStyleSetter(mf), getCellStyleSetter(ma)))

    pub def flatMap(k: a -> CellStyleSetter[r, ka, ef1, b] \ ef2, 
                    x: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1 + ef2, b] = 
        CellStyleSetter(Graded.MutStateError.flatMap(k >> getCellStyleSetter, getCellStyleSetter(x)))


    pub def throwError(msg: String): CellStyleSetter[r, ka, ef, a] = 
        CellStyleSetter(Graded.MutStateError.throwError(msg))

    pub def mapError(f: String -> String \ ef1, 
                        ma: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef + ef1, a] =
        CellStyleSetter(Graded.MutStateError.mapError(f, getCellStyleSetter(ma)))

    pub def catchError(ma: CellStyleSetter[r, ka, ef, a], 
                        hdlr: String -> CellStyleSetter[r, ka, ef1, a] \ ef2): CellStyleSetter[r, ka, ef + ef1 + ef2, a] =
        CellStyleSetter(Graded.MutStateError.catchError(getCellStyleSetter(ma), hdlr >> getCellStyleSetter))

    pub def reifyError(ma: CellStyleSetter[r, ka, ef, a]): CellStyleSetter[r, ka, ef, Result[String, a]] =
        CellStyleSetter(Graded.MutStateError.reifyError(getCellStyleSetter(ma)))


    def liftSetter!(f: CellStyle[r] -> a \ ef): CellStyleSetter[r, ka, ef, a] =
        CellStyleSetter(Graded.MutStateError.liftSetter!(f))
/*
    // def liftSetterResult!(f: CellStyle[r] -> Result[a, err] \ ef): CellStyleSetter[r, ka, a] \ ef with ToString[err] =
    //     CellStyleSetter(Graded.MutStateError.liftSetterResult!(f))
    
    pub def liftAction(f: Unit -> a \ ef): CellStyleSetter[r, ka, a] \ ef =
        CellStyleSetter(Graded.MutStateError.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): CellStyleSetter[r, ka, a] \ ef with ToString[err] =
        CellStyleSetter(Graded.MutStateError.liftActionResult(f))
*/

    pub def setFont!(font: Font[r]): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setFont!(font))

    
    pub def setWrapText!(wrapped: Bool): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setWrapText!(wrapped))


    pub def setIndention!(indent: Int16): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setIndention!(indent))

    pub def setFillForegroundColor!(color: Int16): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setFillForegroundColor!(color))
  
    pub def setFillBackgroundColor!(color: Int16): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setFillBackgroundColor!(color))


    pub def setShrinkToFit!(shrink: Bool): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setShrinkToFit!(shrink))

    pub def setDataFormat!(value: Int16): CellStyleSetter[r, ka, r, Unit] =
        liftSetter!(Interop.Apache.Poi.SS.Usermodel.CellStyle.setDataFormat!(value))


    // pub def cloneStyleFrom!(x: CellStyle[r]): CellStyleSetter[r, ka, Unit] \ r =
    //     let CellStyle(x1) = x;
    //     liftPrim(r -> SheetIO.Internal.CellStyle.cloneStyleFrom!(r, x1))



}
