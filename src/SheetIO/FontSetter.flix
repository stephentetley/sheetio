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

    pub opaque type FontSetter[a] = ##org.apache.poi.ss.usermodel.Font -> Result[a, String] & Impure

    instance Functor[FontSetter] {
        pub def map(f: a -> b & ef, o: FontSetter[a]): FontSetter[b] & ef = 
            SheetIO/FontSetter.map(f, o)
    }

    instance Applicative[FontSetter] {
        pub def point(x: a): FontSetter[a] = 
            SheetIO/FontSetter.point(x)

        pub def ap(f: FontSetter[a -> b & ef], x: FontSetter[a]) : FontSetter[b] & ef = 
            SheetIO/FontSetter.ap(f, x)
    }

    instance Monad[FontSetter] {
        pub def flatMap(f: a -> FontSetter[b] & ef, o: FontSetter[a]): FontSetter[b] & ef =
            SheetIO/FontSetter.flatMap(f, o)
    }


}

namespace SheetIO/FontSetter {

    use SheetIO.FontSetter;
    use SheetIO.FontSetter.FontSetter;

    pub def point(x: a): FontSetter[a] =
        FontSetter(_ -> Ok(x) as & Impure)

    def apply1(x: FontSetter[a], r: ##org.apache.poi.ss.usermodel.Font): Result[a, String] & Impure =
        let FontSetter(f) = x;
        f(r)

    pub def throwError(errMsg: String): FontSetter[a] =
        FontSetter(_ -> Err(errMsg) as & Impure)

    pub def mapError(f: String -> String, ma: FontSetter[a]): FontSetter[a] =
        FontSetter(r -> match apply1(ma, r) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => Err(f(e))
        })


    pub def map(f: a -> b & ef, x: FontSetter[a]): FontSetter[b] & ef =
        FontSetter(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: FontSetter[a -> b & ef], ma: FontSetter[a]): FontSetter[b] & ef =
        FontSetter(r ->
            match apply1(mf, r) {
                case Ok(f) => {
                    match apply1(ma, r) {
                        case Ok(a) => Ok(f(a) as & Impure)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    

    pub def bind(x: FontSetter[a], k: a -> FontSetter[b] & ef): FontSetter[b] & ef =
        FontSetter(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a) as & Impure; apply1(ans, r) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(k: a -> FontSetter[b] & ef, x: FontSetter[a]): FontSetter[b] & ef = bind(x, k)

    def liftPrim(update: ##org.apache.poi.ss.usermodel.Font -> Result[a, String] & Impure): FontSetter[a] =
        FontSetter(r -> update(r))

    pub def liftAction(fn: Unit -> a & Impure): FontSetter[a] =
        FontSetter(_ -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): FontSetter[a] =
        FontSetter(_ -> fn())

    pub def liftOption(x: Option[a], errMsg: String): FontSetter[a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }


    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): FontSetter[a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): FontSetter[a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }




    pub def setFontName!(x: String): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setFontName(r, x))


    pub def setFontHeight!(x: Int16): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setFontHeight(r, x))

    pub def setFontHeightInPoints!(x: Int16): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setFontHeightInPoints(r, x))

    pub def setBold!(x: Bool): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setBold(r, x))

    pub def setItalic!(x: Bool): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setItalic(r, x))

    pub def setStrikeout!(x: Bool): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setStrikeout(r, x))
    
    pub def setColor!(x: Int16): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setColor(r, x))

    /// TODO others need enums...

}