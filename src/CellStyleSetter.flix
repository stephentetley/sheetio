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

    pub opaque type CellStyleSetter[a] = ##org.apache.poi.ss.usermodel.CellStyle -> Result[a, String] & Impure


    instance Functor[CellStyleSetter] {
        pub def map(f: a -> b & ef, o: CellStyleSetter[a]): CellStyleSetter[b] & ef = 
            SheetIO/CellStyleSetter.map(f, o)
    }

    instance Applicative[CellStyleSetter] {
        pub def point(x: a): CellStyleSetter[a] = 
            SheetIO/CellStyleSetter.point(x)

        pub def ap(f: CellStyleSetter[a -> b & ef], x: CellStyleSetter[a]) : CellStyleSetter[b] & ef = 
            SheetIO/CellStyleSetter.ap(f, x)
    }

    instance Monad[CellStyleSetter] {
        pub def flatMap(f: a -> CellStyleSetter[b] & ef, o: CellStyleSetter[a]): CellStyleSetter[b] & ef =
            SheetIO/CellStyleSetter.flatMap(f, o)
    }

}

namespace SheetIO/CellStyleSetter {

    use SheetIO.CellStyleSetter;
    use SheetIO.CellStyleSetter.{CellStyleSetter};
    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.Font;
    use SheetIO.Font.Font;
    

    pub def point(x: a): CellStyleSetter[a] =
        CellStyleSetter(_ -> Ok(x) as & Impure)

    def apply1(x: CellStyleSetter[a], r: ##org.apache.poi.ss.usermodel.CellStyle): Result[a, String] & Impure =
        let CellStyleSetter(f) = x;
        f(r)

    pub def throwError(errMsg: String): CellStyleSetter[a] =
        CellStyleSetter(_ -> Err(errMsg) as & Impure)

     pub def mapError(f: String -> String, ma: CellStyleSetter[a]): CellStyleSetter[a] =
        CellStyleSetter(r -> match apply1(ma, r) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => Err(f(e))
        })


    pub def map(f: a -> b & ef, x: CellStyleSetter[a]): CellStyleSetter[b] & ef =
        CellStyleSetter(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: CellStyleSetter[a -> b & ef], ma: CellStyleSetter[a]): CellStyleSetter[b] & ef =
        CellStyleSetter(r ->
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

    pub def bind(x: CellStyleSetter[a], k: a -> CellStyleSetter[b] & ef): CellStyleSetter[b] & ef =
        CellStyleSetter(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a) as & Impure; apply1(ans, r) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(k: a -> CellStyleSetter[b] & ef, x: CellStyleSetter[a]): CellStyleSetter[b] & ef = 
        bind(x, k)

    def liftPrim(update: ##org.apache.poi.ss.usermodel.CellStyle -> Result[a, String] & Impure): CellStyleSetter[a] =
        CellStyleSetter(r -> update(r))

    pub def liftAction(fn: Unit -> a & Impure): CellStyleSetter[a] =
        CellStyleSetter(_ -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): CellStyleSetter[a] =
        CellStyleSetter(_ -> fn())

    pub def liftOption(x: Option[a], errMsg: String): CellStyleSetter[a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): CellStyleSetter[a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): CellStyleSetter[a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
    }


    pub def setFont!(font: Font): CellStyleSetter[Unit] =
        liftPrim(r -> { 
            let Font(font1) = font; 
            SheetIO/Internal/CellStyle.setFont(r, font1)
        })

    
    pub def setWrapText!(x: Bool): CellStyleSetter[Unit] =
        liftPrim(r -> SheetIO/Internal/CellStyle.setWrapText(r, x))
    
    pub def setIndention!(x: Int16): CellStyleSetter[Unit] =
        liftPrim(r -> SheetIO/Internal/CellStyle.setIndention(r, x))

    pub def setFillForegroundColor!(x: Int16): CellStyleSetter[Unit] =
        liftPrim(r -> SheetIO/Internal/CellStyle.setFillForegroundColor(r, x))
    
    pub def setFillBackgroundColor!(x: Int16): CellStyleSetter[Unit] =
        liftPrim(r -> SheetIO/Internal/CellStyle.setFillBackgroundColor(r, x))

    pub def setShrinkToFit!(x: Bool): CellStyleSetter[Unit] =
        liftPrim(r -> SheetIO/Internal/CellStyle.setShrinkToFit(r, x))

    pub def setDataFormat!(x: Int16): CellStyleSetter[Unit] =
        liftPrim(r -> SheetIO/Internal/CellStyle.setDataFormat(r, x))

    pub def cloneStyleFrom!(x: CellStyle): CellStyleSetter[Unit] =
        let CellStyle(x1) = x;
        liftPrim(r -> SheetIO/Internal/CellStyle.cloneStyleFrom(r, x1))

}