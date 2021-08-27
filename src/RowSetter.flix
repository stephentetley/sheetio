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

    pub opaque type RowSetter[a] = ##org.apache.poi.ss.usermodel.Row -> Option[CellStyle] ~> Result[a, String]

}

namespace SheetIO/RowSetter {

    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.RowSetter;
    use SheetIO.RowSetter.{RowSetter};

    pub def return(x: a): RowSetter[a] =
        RowSetter((_, _) -> Ok(x) as & Impure)

    def apply1(x: RowSetter[a], r: ##org.apache.poi.ss.usermodel.Row, style: Option[CellStyle]): Result[a, String] & Impure =
        let RowSetter(f) = x;
        f(r, style)

    pub def throwError(errMsg: String): RowSetter[a] =
        RowSetter((_, _) -> Err(errMsg) as & Impure)

    pub def mapError(f: String -> String, ma: RowSetter[a]): RowSetter[a] =
        RowSetter((r, style) -> match apply1(ma, r, style) { 
            case Ok(a) => Ok(a)
            case Err(e) => Err(f(e))
        })

    pub def swapError(ma: RowSetter[a], errMsg: String): RowSetter[a] =
        RowSetter((r, style)-> match apply1(ma, r, style) { 
            case Ok(a) => Ok(a)
            case Err(_) => Err(errMsg)
        })


    pub def map(f: a -> b & ef, x: RowSetter[a]): RowSetter[b] =
        RowSetter((r, style) -> match apply1(x, r, style) {
            case Ok(a) => {
                let ans = f(a) as & Impure;
                Ok(ans)
            }
            case Err(e) => Err(e)
        })

    pub def ap(mf: RowSetter[a -> b & ef], ma: RowSetter[a]): RowSetter[b] =
        RowSetter((r, style) ->  match apply1(mf, r, style) {
            case Ok(f) => {
                match apply1(ma, r, style) {
                    case Ok(a) => Ok(f(a) as & Impure)
                    case Err(e2) => Err(e2)
                }
            }
            case Err(e1) => Err(e1)
        })

    pub def seqLeft(ma: RowSetter[a], mb: RowSetter[b]): RowSetter[a] =
        RowSetter((r, style) ->  match apply1(ma, r, style) {
            case Ok(a) => match apply1(mb, r, style) {
                case Ok(_) => Ok(a)
                case Err(e2) => Err(e2)
            }
            case Err(e1) => Err(e1)
        })

    pub def seqRight(ma: RowSetter[a], mb: RowSetter[b]): RowSetter[b] =
        RowSetter((r, style) ->  match apply1(ma, r, style) {
            case Ok(_) => apply1(mb, r, style)
            case Err(e1) => Err(e1)
        })


    pub def bind(x: RowSetter[a], k: a -> RowSetter[b]): RowSetter[b] =
        RowSetter((r, style) -> 
            match apply1(x, r, style) {
                case Ok(a) => {let ans = k(a); apply1(ans, r, style)}
                case Err(e) => Err(e)
            })

    pub def flatMap(k: a -> RowSetter[b], x: RowSetter[a]): RowSetter[b] = bind(x, k)


    pub def kleisliLeft(f1: a -> RowSetter[b], f2: b -> RowSetter[c], x: a): RowSetter[c] =
        RowSetter((r, style) ->  
            match apply1(f1(x), r, style) { 
                case Ok(b) => apply1(f2(b), r, style)
                case Err(e) => Err(e)
            }
        )

    pub def kleisliRight(f1: b -> RowSetter[c], f2: a -> RowSetter[b], x: a): RowSetter[c] =
        kleisliLeft(f2, f1, x)


    def liftPrim(fn: ##org.apache.poi.ss.usermodel.Row -> Option[CellStyle] ~> Result[a, String]): RowSetter[a] =
        RowSetter((r, style) -> fn(r, style))

    pub def liftAction(fn: Unit ~> a): RowSetter[a] =
        RowSetter((_, _) -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit ~> Result[a, String]): RowSetter[a] =
        RowSetter((_, _) -> fn())

    pub def liftOption(x: Option[a], errMsg: String): RowSetter[a] = match x { 
        case Some(x1) => return(x1)
        case None => throwError(errMsg)
    }
    
    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): RowSetter[a] = match x { 
        case Ok(x1) => return(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): RowSetter[a] = match x { 
        case Success(x1) => return(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
    }


    def createOrGetCell(row: ##org.apache.poi.ss.usermodel.Row, 
                        style: Option[CellStyle],
                        ix: Int32): Result[##org.apache.poi.ss.usermodel.Cell, String] & Impure = 
        use Result.{flatMap};
        use SheetIO/Internal/Row.{getCell, createCell};
        let* cell   = match getCell(row, ix) {
                        case Ok(a) => if (not Object.isNull(a)) Ok(a) else createCell(row, ix)
                        case Err(_) => createCell(row, ix)
                    };
        let* _      = match style { 
                        case Some(x) => {let CellStyle(style1) = x; SheetIO/Internal/Cell.setCellStyle(cell, style1)}
                        case None => Ok() as & Impure
                    };
        Ok(cell)

    pub def localCellStyle(style: CellStyle, mx: RowSetter[a]): RowSetter[a] = 
        RowSetter((r, _) -> apply1(mx, r, Some(style)))
    
    pub def putBlankAt!(ix: Int32): RowSetter[Unit] =
        use SheetIO.CellType.{TyBlank};
        use Result.flatMap;
        let step = (row, style) -> {
            let* cell   = createOrGetCell(row, style, ix);
            let typ1    = SheetIO/CellType.toNative(TyBlank);
            let* _      = SheetIO/Internal/Cell.setCellType(cell, typ1);
            Ok()
        };
        liftPrim(step)


    pub def putStringAt!(ix: Int32, value: String): RowSetter[Unit] =
        use Result.flatMap;
        let step = (row, style) -> {
            let* cell   = createOrGetCell(row, style, ix);
            let* _      = SheetIO/Internal/Cell.setStringCellValue(cell, value);
            Ok()
        };
        liftPrim(step)

    pub def putBoolAt!(ix: Int32, value: Bool): RowSetter[Unit] =
        use Result.flatMap;
        let step = (row, style) -> {
            let* cell   = createOrGetCell(row, style, ix);
            let* _      = SheetIO/Internal/Cell.setBooleanCellValue(cell, value);
            Ok()
        };
        liftPrim(step)

    def putNumericAt!(ix: Int32, errMsg: String, value: Float64): RowSetter[Unit] =
        use Result.flatMap;
        let step = (row, style) -> {
            let* cell   = createOrGetCell(row, style, ix);
            let* _      = SheetIO/Internal/Cell.setNumericCellValue(cell, value);
            Ok()
        };
        liftPrim(step) `swapError` errMsg

    pub def putInt8At!(ix: Int32, n: Int8): RowSetter[Unit] =
        putNumericAt!(ix, "putInt8At!", Int8.toFloat64(n))

    pub def putInt16At!(ix: Int32, n: Int16): RowSetter[Unit] =
        putNumericAt!(ix, "putInt16At!", Int16.toFloat64(n))

    pub def putInt32At!(ix: Int32, n: Int32): RowSetter[Unit] =
        putNumericAt!(ix, "putInt32At!", Int32.toFloat64(n))

    pub def putInt64At!(ix: Int32, n: Int64): RowSetter[Unit] =
        putNumericAt!(ix, "putInt64At!", Int64.toFloat64(n))

    pub def putFloat32At!(ix: Int32, n: Float32): RowSetter[Unit] =
        putNumericAt!(ix, "putFloat32At!", Float32.toFloat64(n))

    pub def putFloat64At!(ix: Int32, n: Float64): RowSetter[Unit] =
        putNumericAt!(ix, "putFloat64At!", n)


    pub def putLocalDateTimeAt!(ix: Int32, dt: Time.LocalDateTime): RowSetter[Unit] =
        use Time.LocalDateTime.{LocalDateTime};
        use Result.flatMap;
        let step = (row, style) -> {
            let* cell               = createOrGetCell(row, style, ix);
            let  LocalDateTime(dt1) = dt;
            let* _                  = SheetIO/Internal/Cell.setLocalDateTimeCellValue(cell, dt1);
            Ok()
        };
        liftPrim(step)

    pub def putLocalDateAt!(ix: Int32, dt: Time.LocalDate): RowSetter[Unit] =
        use Time.LocalDate.{LocalDate};
        use Result.flatMap;
        let step = (row, style) -> {
            let* cell           = createOrGetCell(row, style, ix);
            let  LocalDate(dt1) = dt;
            let* _              = SheetIO/Internal/Cell.setLocalDateCellValue(cell, dt1);
            Ok()
        };
        liftPrim(step)


    pub def putHeaders!(arr: Array[String]): RowSetter[Unit] = 
        putHeadersHelper!(arr, Array.length(arr), 0)
        
    def putHeadersHelper!(arr: Array[String], end: Int32, ix: Int32): RowSetter[Unit] =
        if (ix >= end)
            return()
        else {
            let* s1 = liftAction(_ -> arr[ix]);
            let* _  = putStringAt!(ix, s1);
            putHeadersHelper!(arr, end, ix + 1)
        }

    pub def putOptionStringAt!(ix: Int32, x: Option[String]): RowSetter[Unit] = match x {
        case Some(s) => putStringAt!(ix, s)
        case None => putBlankAt!(ix)
    }

    pub def putOptionBoolAt!(ix: Int32, x: Option[Bool]): RowSetter[Unit] = match x {
        case Some(b) => putBoolAt!(ix, b)
        case None => putBlankAt!(ix)
    }

    pub def putOptionInt8At!(ix: Int32, x: Option[Int8]): RowSetter[Unit] = match x {
        case Some(n) => putInt8At!(ix, n)
        case None => putBlankAt!(ix)
        }

    pub def putOptionInt16At!(ix: Int32, x: Option[Int16]): RowSetter[Unit] = match x {
        case Some(n) => putInt16At!(ix, n)
        case None => putBlankAt!(ix)
    }

    pub def putOptionInt32At!(ix: Int32, x: Option[Int32]): RowSetter[Unit] = match x {
        case Some(n) => putInt32At!(ix, n)
        case None => putBlankAt!(ix)
    }

    pub def putOptionInt64At!(ix: Int32, x: Option[Int64]): RowSetter[Unit] = match x {
        case Some(n) => putInt64At!(ix, n)
        case None => putBlankAt!(ix)
    }

    pub def putOptionFloat32At!(ix: Int32, x: Option[Float32]): RowSetter[Unit] = match x {
        case Some(n) => putFloat32At!(ix, n)
        case None => putBlankAt!(ix)
    }

    pub def putOptionFloat64At!(ix: Int32, x: Option[Float64]): RowSetter[Unit] = match x {
        case Some(n) => putFloat64At!(ix, n)
        case None => putBlankAt!(ix)
    }


    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, 
                  o1: RowSetter[t1], o2: RowSetter[t2]): RowSetter[u] =
        ap(map(f, o1), o2)

    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3]): RowSetter[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4]): RowSetter[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5]): RowSetter[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5],
                  o6: RowSetter[t6]): RowSetter[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5],
                  o6: RowSetter[t6],
                  o7: RowSetter[t7]): RowSetter[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef,
                      o1: RowSetter[t1],
                      o2: RowSetter[t2],
                      o3: RowSetter[t3],
                      o4: RowSetter[t4],
                      o5: RowSetter[t5],
                      o6: RowSetter[t6],
                      o7: RowSetter[t7],
                      o8: RowSetter[t8]): RowSetter[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5],
                  o6: RowSetter[t6],
                  o7: RowSetter[t7],
                  o8: RowSetter[t8],
                  o9: RowSetter[t9]): RowSetter[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef,
                   o1: RowSetter[t1],
                   o2: RowSetter[t2],
                   o3: RowSetter[t3],
                   o4: RowSetter[t4],
                   o5: RowSetter[t5],
                   o6: RowSetter[t6],
                   o7: RowSetter[t7],
                   o8: RowSetter[t8],
                   o9: RowSetter[t9],
                   o10: RowSetter[t10]): RowSetter[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)



    /// Alias for `seqRight`
    pub def *>(p1: RowSetter[a], p2: RowSetter[b]): RowSetter[b] = seqRight(p1, p2)

    /// Alias for `seqLeft`
    pub def <*(p1: RowSetter[a], p2: RowSetter[b]): RowSetter[a] = seqLeft(p1, p2)


    // Haskell's <$>
    pub def <&>(f: a -> b, p2: RowSetter[a]): RowSetter[b] =
        map(f, p2)

    pub def <&(x: b, p2: RowSetter[a]): RowSetter[b] =
        p2 *> return(x)

    /// Alias for `bind`
    pub def >>=(ma: RowSetter[a], k: a -> RowSetter[b]): RowSetter[b] = 
        bind(ma, k)

    /// Alias for `flatMap` (flipped `bind`)
    pub def =<<(k: a -> RowSetter[b], ma: RowSetter[a]): RowSetter[b] = 
        flatMap(k, ma)


    // Haskell's ap <*>
    pub def <*>(mf: RowSetter[a -> b], p2: RowSetter[a]): RowSetter[b] =
        ap(mf, p2)

    /// Alias for `kleisliLeft`
    pub def >=>(f1: a -> RowSetter[b], f2: b -> RowSetter[c]): a -> RowSetter[c] = 
        x -> kleisliLeft(f1, f2, x)

    /// Alias for `kleisliRight`
    pub def <=<(f1: b -> RowSetter[c], f2: a -> RowSetter[b]): a -> RowSetter[c] = 
        x -> kleisliRight(f1, f2, x)


    pub def mapM(mf: a -> RowSetter[b], xs: List[a]): RowSetter[List[b]] =
        RowSetter((r, style) ->
            mapMHelper(r, style, mf, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(row: ##org.apache.poi.ss.usermodel.Row, 
                   style: Option[CellStyle], 
                   mf: a -> RowSetter[b], 
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

    pub def mapX(mf: a -> RowSetter[b], xs: List[a]): RowSetter[Unit] =
        RowSetter((r, style) ->
            mapXHelper(r, style, mf, xs, e -> Err(e), _ -> Ok())
        )

    def mapXHelper(row: ##org.apache.poi.ss.usermodel.Row, 
                   style: Option[CellStyle],
                   mf: a -> RowSetter[b], 
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

}

