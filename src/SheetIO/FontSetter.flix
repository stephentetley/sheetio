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

    use SheetIO/Internal/PrimError.{PrimError};

    pub enum FontSetter[r: Region, a: Type](##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] & r)

    instance Functor[FontSetter[r]] {
        pub def map(f: a -> b & ef, o: FontSetter[r, a]): FontSetter[r, b] & ef = 
            SheetIO/FontSetter.map(f, o)
    }

    instance Applicative[FontSetter[r]] {
        pub def point(x: a): FontSetter[r, a] = 
            SheetIO/FontSetter.point(x)

        pub def ap(f: FontSetter[r, a -> b & ef], x: FontSetter[r, a]) : FontSetter[r, b] & ef = 
            SheetIO/FontSetter.ap(f, x)
    }

    instance Monad[FontSetter[r]] {
        pub def flatMap(f: a -> FontSetter[r, b] & ef, o: FontSetter[r, a]): FontSetter[r, b] & ef =
            SheetIO/FontSetter.flatMap(f, o)
    }


}

namespace SheetIO/FontSetter {

    use SheetIO.FontSetter;
    use SheetIO.FontSetter.FontSetter;

    use SheetIO/Internal/PrimError.{PrimError};

    pub def point(x: a): FontSetter[r, a] =
        FontSetter(_ -> Ok(x) as & r)

    def apply1(x: FontSetter[r, a], r: ##org.apache.poi.ss.usermodel.Font): Result[a, PrimError] & r =
        let FontSetter(f) = x;
        f(r)

    def getFontSetterRegion(): FontSetter[r, Region[r]] = 
        FontSetter(_ -> Ok(() as Region[r]) as & r)

    pub def throwError(msg: String): FontSetter[r, a] =
        FontSetter(_ -> Err({source = msg, exnMessage = ""}) as & r)

    pub def mapError(f: String -> String, ma: FontSetter[r, a]): FontSetter[r, a] =
        FontSetter(s -> match apply1(ma, s) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => let src = e.source; Err({source = f(src) | e})
        })


    pub def map(f: a -> b & ef, x: FontSetter[r, a]): FontSetter[r, b] & ef =
        FontSetter(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & r;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: FontSetter[r, a -> b & ef], ma: FontSetter[r, a]): FontSetter[r, b] & ef =
        FontSetter(r ->
            match apply1(mf, r) {
                case Ok(f) => {
                    match apply1(ma, r) {
                        case Ok(a) => Ok(f(a) as & r)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    

    pub def flatMap(k: a -> FontSetter[r, b] & ef, x: FontSetter[r, a]): FontSetter[r, b] & ef = 
        FontSetter(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a) as & r; apply1(ans, r) }
                case Err(e) => Err(e)
            }
        ) as & ef


    pub def bind(x: FontSetter[r, a], k: a -> FontSetter[r, b] & ef): FontSetter[r, b] & ef = flatMap(k, x)

    def liftPrim(update: ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] & ef): FontSetter[r, a] & ef=
        FontSetter(r -> update(r) as & r) as & ef

    pub def liftAction(fn: Unit -> a & ef): FontSetter[r, a] & ef =
        FontSetter(_ -> let x = fn() as & r; Ok(x)) as & ef

    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): FontSetter[r, a] & ef =
        FontSetter(_ -> let ans = fn() as & r; Result.mapErr(s -> {source = s, exnMessage = ""}, ans)) as & ef

    pub def liftOption(errMsg: String, x: Option[a]): FontSetter[r, a] = match x { 
        case Some(x1) => point(x1)
        case None     => throwError(errMsg)
    }


    pub def liftResult(x: Result[a, String]): FontSetter[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(msg) => throwError(msg)
    }


    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): FontSetter[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): FontSetter[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    pub def setFontName!(x: String): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setFontName!(r, x))


    pub def setFontHeight!(x: Int16): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setFontHeight!(r, x))

    pub def setFontHeightInPoints!(x: Int16): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setFontHeightInPoints!(r, x))

    pub def setBold!(x: Bool): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setBold!(r, x))

    pub def setItalic!(x: Bool): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setItalic!(r, x))

    pub def setStrikeout!(x: Bool): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setStrikeout!(r, x))
    
    pub def setColor!(x: Int16): FontSetter[r, Unit] \ Write(r) = 
        let* r = getFontSetterRegion();
        liftPrim(SheetIO/Internal/Font.setColor!(r, x))

    /// TODO others need enums...

}
