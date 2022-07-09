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

    pub enum CellStyleSetter[r: Region, a: Type](##org.apache.poi.ss.usermodel.CellStyle -> Result[a, PrimError] & r)


    instance Functor[CellStyleSetter[r]] {
        pub def map(f: a -> b & ef, o: CellStyleSetter[r, a]): CellStyleSetter[r, b] & ef = 
            SheetIO/CellStyleSetter.map(f, o)
    }

    instance Applicative[CellStyleSetter[r]] {
        pub def point(x: a): CellStyleSetter[r, a] = 
            SheetIO/CellStyleSetter.point(x)

        pub def ap(f: CellStyleSetter[r, a -> b & ef], x: CellStyleSetter[r, a]) : CellStyleSetter[r, b] & ef = 
            SheetIO/CellStyleSetter.ap(f, x)
    }

    instance Monad[CellStyleSetter[r]] {
        pub def flatMap(f: a -> CellStyleSetter[r, b] & ef, o: CellStyleSetter[r, a]): CellStyleSetter[r, b] & ef =
            SheetIO/CellStyleSetter.flatMap(f, o)
    }

}

namespace SheetIO/CellStyleSetter {

    use SheetIO.CellStyleSetter;
    use SheetIO.CellStyleSetter.{CellStyleSetter};

    use SheetIO/Internal/PrimError.{PrimError};
    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.Font;
    use SheetIO.Font.Font;
    

    pub def point(x: a): CellStyleSetter[r, a] =
        CellStyleSetter(_ -> Ok(x) as & r)

    def apply1(x: CellStyleSetter[r, a], r: ##org.apache.poi.ss.usermodel.CellStyle): Result[a, PrimError] & r =
        let CellStyleSetter(f) = x;
        f(r)

    def getCellStyleSetterRegion(): CellStyleSetter[r, Region[r]] = 
        CellStyleSetter(_ -> Ok(() as Region[r]) as & r)


    pub def throwError(msg: String): CellStyleSetter[r, a] =
        CellStyleSetter(_ -> Err({source = msg, exnMessage = ""}) as & r)

     pub def mapError(f: String -> String, ma: CellStyleSetter[r, a]): CellStyleSetter[r, a] =
        CellStyleSetter(s -> match apply1(ma, s) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => let src = e.source; Err({source = f(src) | e})
        })


    pub def map(f: a -> b & ef, x: CellStyleSetter[r, a]): CellStyleSetter[r, b] & ef =
        CellStyleSetter(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & r;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: CellStyleSetter[r, a -> b & ef], ma: CellStyleSetter[r, a]): CellStyleSetter[r, b] & ef =
        CellStyleSetter(r ->
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

    
    pub def flatMap(k: a -> CellStyleSetter[r, b] & ef, x: CellStyleSetter[r, a]): CellStyleSetter[r, b] & ef = 
        CellStyleSetter(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a) as & r; apply1(ans, r) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def bind(x: CellStyleSetter[r, a], k: a -> CellStyleSetter[r, b] & ef): CellStyleSetter[r, b] & ef =
        flatMap(k, x)

    def liftPrim(update: ##org.apache.poi.ss.usermodel.CellStyle -> Result[a, PrimError] & ef): CellStyleSetter[r, a] & ef=
        CellStyleSetter(r -> update(r) as & r) as & ef

    pub def liftAction(fn: Unit -> a & ef): CellStyleSetter[r, a] & ef=
        CellStyleSetter(_ -> let x = fn() as & r; Ok(x)) as & ef

    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): CellStyleSetter[r, a] & ef=
        CellStyleSetter(_ -> let ans = fn() as & r; Result.mapErr(s -> {source = s, exnMessage = ""}, ans)) as & ef

    pub def liftOption(x: Option[a], errMsg: String): CellStyleSetter[r, a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): CellStyleSetter[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): CellStyleSetter[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): CellStyleSetter[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    pub def setFont!(font: Font): CellStyleSetter[r, Unit] \ Write(r) =    
        let Font(font1) = font; 
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setFont!(r, font1))


    pub def setWrapText!(x: Bool): CellStyleSetter[r, Unit] \ Write(r) =
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setWrapText!(r, x))


    pub def setIndention!(x: Int16): CellStyleSetter[r, Unit] \ Write(r) =
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setIndention!(r, x))

    pub def setFillForegroundColor!(x: Int16): CellStyleSetter[r, Unit] \ Write(r) =
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setFillForegroundColor!(r, x))
    
    pub def setFillBackgroundColor!(x: Int16): CellStyleSetter[r, Unit] \ Write(r) =
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setFillBackgroundColor!(r, x))

    pub def setShrinkToFit!(x: Bool): CellStyleSetter[r, Unit] \ Write(r) =
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setShrinkToFit!(r, x))

    pub def setDataFormat!(x: Int16): CellStyleSetter[r, Unit] \ Write(r) =
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.setDataFormat!(r, x))

    pub def cloneStyleFrom!(x: CellStyle): CellStyleSetter[r, Unit] \ Write(r) =
        let CellStyle(x1) = x;
        let* r = getCellStyleSetterRegion();
        liftPrim(SheetIO/Internal/CellStyle.cloneStyleFrom!(r, x1))

}
