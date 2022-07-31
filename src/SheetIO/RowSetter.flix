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
    use SheetIO/Internal/EvalError.{EvalError};

    pub enum RowSetter[r: Region, a: Type](
        Option[CellStyle[r]] -> ##org.apache.poi.ss.usermodel.Row -> Result[a, EvalError] & r)

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

    instance MonadZero[RowSetter[r]] {
        pub def empty(): RowSetter[r, a] = SheetIO/RowSetter.throwError("empty")
    }

    // No MonadOr

}

namespace SheetIO/RowSetter {

    use SheetIO.RowSetter;
    use SheetIO.RowSetter.{RowSetter};
    
    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO/Internal/PrimError.{PrimError};

    pub def point(x: a): RowSetter[r, a] =
        RowSetter((_, _) -> Ok(x) as & r)

    def apply1(x: RowSetter[r, a], style: Option[CellStyle[r]], row: ##org.apache.poi.ss.usermodel.Row): Result[a, EvalError] & r =
        let RowSetter(f) = x;
        f(style, row)




    pub def throwError(msg: String): RowSetter[r, a] =
        RowSetter((_, row) -> 
            Err(SheetIO/Internal/EvalError.create(() as Region[r], row, {source = msg, exnMessage = ""}))
        )



    pub def swapError(errSrc: String, ma: RowSetter[r, a]): RowSetter[r, a] =
        RowSetter((style, row) -> match apply1(ma, style, row) { 
            case Ok(x1)   => Ok(x1)
            case Err(err) => Err({source = errSrc | err})
        })


    pub def map(f: a -> b & ef, x: RowSetter[r, a]): RowSetter[r, b] & ef =
        RowSetter((style, row) -> match apply1(x, style, row) {
            case Ok(a) => {
                let ans = f(a) as & r;
                Ok(ans)
            }
            case Err(e) => Err(e)
        }) as & ef

    pub def ap(mf: RowSetter[r, a -> b & ef], ma: RowSetter[r, a]): RowSetter[r, b] & ef =
        RowSetter((style, row) ->  match apply1(mf, style, row) {
            case Ok(f) => {
                match apply1(ma, style, row) {
                    case Ok(a) => Ok(f(a) as & r)
                    case Err(e2) => Err(e2)
                }
            }
            case Err(e1) => Err(e1)
        }) as & ef


    pub def flatMap(k: a -> RowSetter[r, b] & ef, x: RowSetter[r, a]): RowSetter[r, b] & ef =
        RowSetter((style, row) -> 
            match apply1(x, style, row) {
                case Ok(a) => {let ans = k(a) as & r; apply1(ans, style, row)}
                case Err(e) => Err(e)
            }) as & ef



    def liftPrim(fn: Region[r] -> Option[CellStyle[r]] -> ##org.apache.poi.ss.usermodel.Row -> Result[a, PrimError] & ef): RowSetter[r, a] & ef =
        RowSetter((style, row) -> 
            let r = () as Region[r]; 
            {fn(r, style, row) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.create(r, row))
        ) as & ef
        

    pub def liftAction(fn: Unit -> a & ef): RowSetter[r, a] & ef=
        RowSetter((_, _) -> 
            let x = fn() as & r; 
            Ok(x)
        ) as & ef

    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): RowSetter[r, a] & ef=
        RowSetter((_, row) -> 
            let r = () as Region[r]; 
            {fn() as & r} |> Result.mapErr(s -> SheetIO/Internal/EvalError.create(r, row, {source = s, exnMessage = ""}))
        ) as & ef

    pub def liftOption(x: Option[a], errMsg: String): RowSetter[r, a] = match x { 
        case Some(x1) => point(x1)
        case None     => throwError(errMsg)
    }
    
    pub def liftResult(x: Result[a, String]): RowSetter[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }
    
    pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): RowSetter[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(makeErrMsg: List[e] -> String, x: Validation[a, e]): RowSetter[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }



    /// This sets the CellStyle...
    def createOrGetCell(ix: Int32): RowSetter[r, ##org.apache.poi.ss.usermodel.Cell] \ Write(r) = 
        use SheetIO/Internal/Row.{getCell, createCell!};
        use Result.{flatMap};
        let body = (r, style, row) -> {
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

    def liftPrimCell(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Cell -> Option[CellStyle[r]] -> Result[a, PrimError] & ef): RowSetter[r, a] \ { Write(r), ef } =
        let* cell = createOrGetCell(ix);
        liftPrim((r, style, _) -> fn(r, cell, style))


    pub def localCellStyle(style: CellStyle[r], mx: RowSetter[r, a]): RowSetter[r, a] = 
        RowSetter((_, row) -> apply1(mx, Some(style), row))


    pub def putBlankAt!(ix: Int32): RowSetter[r, Unit] \ Write(r) =
        liftPrimCell(ix, (r, cell, _) -> SheetIO/Internal/Cell.setBlank!(r, cell))
            |> swapError("putBlankAt!")


    pub def putStringAt!(ix: Int32, value: String): RowSetter[r, Unit] \ Write(r) =
        liftPrimCell(ix, (r, cell, _) -> SheetIO/Internal/Cell.setCellValueString!(r, value, cell))
            |> swapError("putBoolAt!")

    pub def putBoolAt!(ix: Int32, value: Bool): RowSetter[r, Unit] \ Write(r) =
        liftPrimCell(ix, (r, cell, _) -> SheetIO/Internal/Cell.setCellValueBoolean!(r, value, cell))
            |> swapError("putBoolAt!")



    def putNumericAt!(ix: Int32, value: Float64): RowSetter[r, Unit] \ Write(r) =
        liftPrimCell(ix, (r, cell, _) -> SheetIO/Internal/Cell.setCellValueDouble!(r, value, cell))


    pub def putInt8At!(ix: Int32, n: Int8): RowSetter[r, Unit] \ Write(r) =
        putNumericAt!(ix, Int8.toFloat64(n))
            |> swapError("putInt8At!")

    pub def putInt16At!(ix: Int32, n: Int16): RowSetter[r, Unit] \ Write(r) =
        putNumericAt!(ix, Int16.toFloat64(n))
            |> swapError("putInt16At")

    pub def putInt32At!(ix: Int32, n: Int32): RowSetter[r, Unit] \ Write(r) =
        putNumericAt!(ix, Int32.toFloat64(n))
            |> swapError("putInt32At!")


    pub def putInt64At!(ix: Int32, n: Int64): RowSetter[r, Unit] \ Write(r) =
        putNumericAt!(ix, Int64.toFloat64(n))
            |> swapError("putInt64At!")

    pub def putFloat32At!(ix: Int32, n: Float32): RowSetter[r, Unit] \ Write(r) =
        putNumericAt!(ix, Float32.toFloat64(n))
            |> swapError("putFloat32At!")

    pub def putFloat64At!(ix: Int32, n: Float64): RowSetter[r, Unit] \ Write(r) =
        putNumericAt!(ix, n)
            |> swapError("putFloat64At!")

    pub def putLocalDateTimeAt!(ix: Int32, dt: Time.LocalDateTime): RowSetter[r, Unit] \ Write(r) =
        use Time.LocalDateTime.{LocalDateTime};
        let  LocalDateTime(dt1) = dt;
        liftPrimCell(ix, (r, cell, _) -> SheetIO/Internal/Cell.setCellValueLocalDateTime!(r, dt1, cell))



    pub def putLocalDateAt!(ix: Int32, dt: Time.LocalDate): RowSetter[r, Unit] \ Write(r) =
        use Time.LocalDate.{LocalDate};
        let  LocalDate(dt1) = dt;
        liftPrimCell(ix, (r, cell, _) -> SheetIO/Internal/Cell.setCellValueLocalDate!(r, dt1, cell))
 
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

}

