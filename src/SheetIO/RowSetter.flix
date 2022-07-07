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


/// Use `Option[CellStyle]` as a Reader env in the Monad to allow inherited styles.

namespace SheetIO {

    use SheetIO.CellStyle;

    pub enum RowSetter[r: Region, a: Type](
        ##org.apache.poi.ss.usermodel.Row -> Option[CellStyle] -> Result[a, String] & r)

    instance Functor[RowSetter[r]] {
        pub def map(f: a -> b & ef, o: RowSetter[r, a]): RowSetter[r, b] & ef = 
            SheetIO/RowSetter.map(f, o)
    }

    instance Applicative[RowSetter[r]] {
        pub def point(x: a): RowSetter[r, a] = 
            SheetIO/RowSetter.point(x)

        pub def ap(f: RowSetter[r, a -> b & ef], x: RowSetter[r, a]): RowSetter[r, b] & ef = 
            SheetIO/RowSetter.ap(f, x)
    }

    instance Monad[RowSetter[r]] {
        pub def flatMap(f: a -> RowSetter[r, b] & ef, o: RowSetter[r, a]): RowSetter[r, b] & ef =
            SheetIO/RowSetter.flatMap(f, o)
    }


}

namespace SheetIO/RowSetter {

    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.RowSetter;
    use SheetIO.RowSetter.{RowSetter};

    pub def point(x: a): RowSetter[r, a] =
        RowSetter((_, _) -> Ok(x) as & r)

    def apply1(x: RowSetter[r, a], row: ##org.apache.poi.ss.usermodel.Row, style: Option[CellStyle]): Result[a, String] & r =
        let RowSetter(f) = x;
        f(row, style)


    /// DESIGN NOTE - a function like this shouldn't be needed in other modules - extend `liftPrim`
    /// to materialize the region.
    /// def getRowSetterRegion(): RowSetter[r, Region[r]] = 
    ///     RowSetter((_, _) -> Ok(() as Region[r]) as & r)

    pub def throwError(errMsg: String): RowSetter[r, a] =
        RowSetter((_, _) -> Err(errMsg) as & r)

    pub def mapError(f: String -> String, ma: RowSetter[r, a]): RowSetter[r, a] =
        RowSetter((row, style) -> match apply1(ma, row, style) { 
            case Ok(a) => Ok(a)
            case Err(e) => Err(f(e))
        })

    pub def swapError(ma: RowSetter[r, a], errMsg: String): RowSetter[r, a] =
        RowSetter((row, style)-> match apply1(ma, row, style) { 
            case Ok(a) => Ok(a)
            case Err(_) => Err(errMsg)
        })


    pub def map(f: a -> b & ef, x: RowSetter[r, a]): RowSetter[r, b] & ef =
        RowSetter((row, style) -> match apply1(x, row, style) {
            case Ok(a) => {
                let ans = f(a) as & r;
                Ok(ans)
            }
            case Err(e) => Err(e)
        }) as & ef

    pub def ap(mf: RowSetter[r, a -> b & ef], ma: RowSetter[r, a]): RowSetter[r, b] & ef =
        RowSetter((row, style) ->  match apply1(mf, row, style) {
            case Ok(f) => {
                match apply1(ma, row, style) {
                    case Ok(a) => Ok(f(a) as & r)
                    case Err(e2) => Err(e2)
                }
            }
            case Err(e1) => Err(e1)
        }) as & ef


    pub def flatMap(k: a -> RowSetter[r, b] & ef, x: RowSetter[r, a]): RowSetter[r, b] & ef =
        RowSetter((row, style) -> 
            match apply1(x, row, style) {
                case Ok(a) => {let ans = k(a) as & r; apply1(ans, row, style)}
                case Err(e) => Err(e)
            }) as & ef


    pub def bind(x: RowSetter[r, a], k: a -> RowSetter[r, b] & ef): RowSetter[r, b] & ef =
        flatMap(k, x)

    def liftPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Row -> Option[CellStyle] -> Result[a, String] & ef): RowSetter[r, a] & ef =
        RowSetter((row, style) -> {let r = () as Region[r]; fn(r, row, style) as & r}) as & ef

    pub def liftAction(fn: Unit -> a & ef): RowSetter[r, a] & ef=
        RowSetter((_, _) -> let x = fn() as & r; Ok(x)) as & ef

    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): RowSetter[r, a] & ef=
        RowSetter((_, _) -> fn() as & r) as & ef

    pub def liftOption(x: Option[a], errMsg: String): RowSetter[r, a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }
    
    pub def liftResult(x: Result[a, String]): RowSetter[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }
    
    pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): RowSetter[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): RowSetter[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    /// def createOrGetCell(r: Region[r], 
    ///                     ix: Int32,
    ///                     style: Option[CellStyle], 
    ///                     row: ##org.apache.poi.ss.usermodel.Row): Result[##org.apache.poi.ss.usermodel.Cell, String] & r = 
    ///     use Result.{flatMap};
    ///     use SheetIO/Internal/Row.{getCell, createCell!};
    ///     let* cell   = match getCell(r, ix, row) {
    ///                     case Ok(a)  => if (not Object.isNull(a)) Ok(a) else createCell!(r, ix, row)
    ///                     case Err(_) => createCell!(r, ix, row)
    ///                 };
    ///     let* _      = match style { 
    ///                     case Some(x) => {let CellStyle(style1) = x; SheetIO/Internal/Cell.setCellStyle!(r, style1, cell)}
    ///                     case None    => Ok() as & r
    ///                 };
    ///     Ok(cell)


    def createOrGetCell(ix: Int32): RowSetter[r, ##org.apache.poi.ss.usermodel.Cell] \ Write(r) = 
        use SheetIO/Internal/Row.{getCell, createCell!};
        use Result.{flatMap};
        let body = (r, row, style) -> {
            let* cell   = match getCell(r, ix, row) {
                            case Ok(a)  => if (not Object.isNull(a)) Ok(a) else createCell!(r, ix, row)
                            case Err(_) => createCell!(r, ix, row)
                        };
            let* _      = match style { 
                            case Some(x) => {let CellStyle(style1) = x; SheetIO/Internal/Cell.setCellStyle!(r, style1, cell)}
                            case None    => Ok() as & r
                        };
            Ok(cell)
        };
        liftPrim(body)

    pub def localCellStyle(style: CellStyle, mx: RowSetter[r, a]): RowSetter[r, a] = 
        RowSetter((row, _) -> apply1(mx, row, Some(style)))


    pub def putBlankAt!(ix: Int32): RowSetter[r, Unit] \ Write(r) =
        let* cell = createOrGetCell(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.setBlank!(r, cell))
        


    pub def putStringAt!(ix: Int32, value: String): RowSetter[r, Unit] \ Write(r) =
        let* cell = createOrGetCell(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.setCellValueString!(r, value, cell))


    pub def putBoolAt!(ix: Int32, value: Bool): RowSetter[r, Unit] \ Write(r) =
        let* cell = createOrGetCell(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.setCellValueBoolean!(r, value, cell))



    def putNumericAt!(ix: Int32, value: Float64): RowSetter[r, Unit] \ Write(r) =
        let* cell = createOrGetCell(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.setCellValueDouble!(r, value, cell))

    pub def putInt8At!(ix: Int32, n: Int8): RowSetter[r, Unit] \ Write(r) =
        mapError(
            String.replaceFirstMatch(regex = "setCellValueDouble", to = "putInt8At"), 
            putNumericAt!(ix, Int8.toFloat64(n))
        )

    pub def putInt16At!(ix: Int32, n: Int16): RowSetter[r, Unit] \ Write(r) =
        mapError(
            String.replaceFirstMatch(regex = "setCellValueDouble", to = "putInt16At"), 
            putNumericAt!(ix, Int16.toFloat64(n))
        )

    pub def putInt32At!(ix: Int32, n: Int32): RowSetter[r, Unit] \ Write(r) =
        mapError(
            String.replaceFirstMatch(regex = "setCellValueDouble", to = "putInt32At"), 
            putNumericAt!(ix, Int32.toFloat64(n))
        )

    pub def putInt64At!(ix: Int32, n: Int64): RowSetter[r, Unit] \ Write(r) =
        mapError(
            String.replaceFirstMatch(regex = "setCellValueDouble", to = "putInt64At"),
            putNumericAt!(ix, Int64.toFloat64(n))
        )

    pub def putFloat32At!(ix: Int32, n: Float32): RowSetter[r, Unit] \ Write(r) =
        mapError(
            String.replaceFirstMatch(regex = "setCellValueDouble", to = "putFloat32At"),
            putNumericAt!(ix, Float32.toFloat64(n))
        )

    pub def putFloat64At!(ix: Int32, n: Float64): RowSetter[r, Unit] \ Write(r) =
        mapError(
            String.replaceFirstMatch(regex = "setCellValueDouble", to = "putFloat64At"),
            putNumericAt!(ix, n)
        )


    pub def putLocalDateTimeAt!(ix: Int32, dt: Time.LocalDateTime): RowSetter[r, Unit] \ Write(r) =
        use Time.LocalDateTime.{LocalDateTime};
        let  LocalDateTime(dt1) = dt;
        let* cell = createOrGetCell(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.setCellValueLocalDateTime!(r, dt1, cell))



    pub def putLocalDateAt!(ix: Int32, dt: Time.LocalDate): RowSetter[r, Unit] \ Write(r) =
        use Time.LocalDate.{LocalDate};
        let  LocalDate(dt1) = dt;
        let* cell = createOrGetCell( ix);
        liftPrim((r, _, _) ->SheetIO/Internal/Cell.setCellValueLocalDate!(r, dt1, cell))
 
    pub def putHeaders!(l: List[String]): RowSetter[r, Unit] \ Write(r) = 
        putHeadersHelper!(l, 0)
        
    def putHeadersHelper!(l: List[String], ix: Int32): RowSetter[r, Unit] \ Write(r) = match l {
        case Nil => point()
        case s :: rs => {
            let* _  = putStringAt!(ix, s);
            putHeadersHelper!(rs, ix + 1)
        }
    }

    pub def putOptionStringAt!(ix: Int32, x: Option[String]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(s) => putStringAt!(ix, s)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionBoolAt!(ix: Int32, x: Option[Bool]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(b) => putBoolAt!(ix, b)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt8At!(ix: Int32, x: Option[Int8]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(n) => putInt8At!(ix, n)
        case None    => putBlankAt!(ix)
        }

    pub def putOptionInt16At!(ix: Int32, x: Option[Int16]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(n) => putInt16At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt32At!(ix: Int32, x: Option[Int32]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(n) => putInt32At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt64At!(ix: Int32, x: Option[Int64]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(n) => putInt64At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionFloat32At!(ix: Int32, x: Option[Float32]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(n) => putFloat32At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionFloat64At!(ix: Int32, x: Option[Float64]): RowSetter[r, Unit] \ Write(r) = match x {
        case Some(n) => putFloat64At!(ix, n)
        case None    => putBlankAt!(ix)
    }

/*

    pub def mapM(mf: a -> RowSetter[r, b], xs: List[a]): RowSetter[r, List[b]] =
        RowSetter((r, style) ->
            mapMHelper(r, style, mf, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(row: ##org.apache.poi.ss.usermodel.Row, 
                   style: Option[CellStyle], 
                   mf: a -> RowSetter[r, b], 
                   xs: List[a], 
                   fk: String -> Result[List[b], String], 
                   sk: List[b] -> Result[List[b], String]): Result[List[b], String] & Impure =
        match xs {
            case Nil => sk(Nil)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, row, style) {
                    case Err(e) => fk(e)
                    case Ok(x1) => mapMHelper(row, style, mf, rs, fk, ks -> sk(x1 :: ks))
                }
            }
        }

    pub def mapX(mf: a -> RowSetter[r, b], xs: List[a]): RowSetter[r, Unit] =
        RowSetter((r, style) ->
            mapXHelper(r, style, mf, xs, e -> Err(e), _ -> Ok())
        )

    def mapXHelper(row: ##org.apache.poi.ss.usermodel.Row, 
                   style: Option[CellStyle],
                   mf: a -> RowSetter[r, b], 
                   xs: List[a], 
                   fk: String -> Result[Unit, String], 
                   sk: Unit -> Result[Unit, String]): Result[Unit, String] & Impure =
        match xs {
            case Nil => sk()
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, row, style) {
                    case Err(e) => fk(e)
                    case Ok(_) => mapXHelper(row, style, mf, rs, fk, sk)
                }
            }
        }

*/
}

