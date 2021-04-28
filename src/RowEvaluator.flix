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

    pub opaque type RowEvaluator[a] = ##org.apache.poi.ss.usermodel.Row ~> Result[a, String]

}

namespace SheetIO/RowEvaluator {

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};

    use SheetIO.CellType;
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};
    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};
    
    pub def return(x: a): RowEvaluator[a] =
        RowEvaluator(_ -> Ok(x) as & Impure)

    def apply1(x: RowEvaluator[a], row: ##org.apache.poi.ss.usermodel.Row): Result[a, String] & Impure =
        let RowEvaluator(f) = x;
        f(row)

    pub def throwError(msg: String): RowEvaluator[a] =
        RowEvaluator(_ -> Err(msg) as & Impure)

    pub def swapError(ma: RowEvaluator[a], msg: String): RowEvaluator[a] =
        RowEvaluator(r ->
            match apply1(ma, r) {
                case Ok(a) => Ok(a)
                case Err(_) => Err(msg)
            }
        )

    pub def map(f: a -> b & ef, x: RowEvaluator[a]): RowEvaluator[b] =
        RowEvaluator(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: RowEvaluator[a -> b & ef], ma: RowEvaluator[a]): RowEvaluator[b] =
        RowEvaluator(r ->
            match apply1(mf, r) {
                case Ok(f) => {
                    match apply1(ma, r) {
                        case Ok(a) => Ok(f(a) as & Impure)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        )
    
    pub def seqLeft(ma: RowEvaluator[a], mb: RowEvaluator[b]): RowEvaluator[a] = 
        RowEvaluator(r ->
            match apply1(ma, r) { 
                case Ok(a) => match apply1(mb, r) {
                    case Ok(_) => Ok(a)
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqRight(ma: RowEvaluator[a], mb: RowEvaluator[b]): RowEvaluator[b] = 
        RowEvaluator(r ->
            match apply1(ma, r) { 
                case Ok(_) => apply1(mb, r)
                case Err(e) => Err(e)
            }
        )

    pub def bind(x: RowEvaluator[a], k: a -> RowEvaluator[b]): RowEvaluator[b] =
        RowEvaluator(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a); apply1(ans, r) }
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(k: a -> RowEvaluator[b], x: RowEvaluator[a]): RowEvaluator[b] = bind(x, k)


    pub def alt(ea: RowEvaluator[a], eb: RowEvaluator[a]): RowEvaluator[a] =
        RowEvaluator(r ->
            match apply1(ea, r) {
                case Ok(a) => Ok(a)
                case Err(_) => apply1(eb, r)
            }
        )

    pub def kleisliLeft(f1: a -> RowEvaluator[b], f2: b -> RowEvaluator[c], x: a): RowEvaluator[c] =
        RowEvaluator(r -> 
            match apply1(f1(x), r) { 
                case Ok(b) => apply1(f2(b), r)
                case Err(e) => Err(e)
            }
        )

    pub def kleisliRight(f1: b -> RowEvaluator[c], f2: a -> RowEvaluator[b], x: a): RowEvaluator[c] =
        kleisliLeft(f2, f1, x)

    pub def optional(x: RowEvaluator[a]): RowEvaluator[Option[a]] =
        RowEvaluator(r -> match apply1(x, r) {
            case Ok(a) => Ok(Some(a))
            case Err(_) => Ok(None)
        })

    /// Note - this returns the default value for any errors (i.e. unknown column which is a
    /// schema error), not just a "read" error if the cell type is wrong.
    pub def withDefault(x: RowEvaluator[a], d: a): RowEvaluator[a] =
        alt(x, return(d))

    def liftPrim(fn: ##org.apache.poi.ss.usermodel.Row ~> Result[a, String]): RowEvaluator[a] =
        RowEvaluator(r -> fn(r))


    pub def liftAction(fn: Unit ~> a): RowEvaluator[a] =
        RowEvaluator(_ -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit ~> Result[a, String]): RowEvaluator[a] =
        RowEvaluator(_ -> fn())


    pub def liftOption(x: Option[a], errMsg: String): RowEvaluator[a] = match x { 
        case Some(x1) => return(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): RowEvaluator[a] = match x { 
        case Ok(x1) => return(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }


    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): RowEvaluator[a] = match x { 
        case Success(x1) => return(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
    }

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    def withCell(row: ##org.apache.poi.ss.usermodel.Row,
                 ix: Int32, 
                 f: ##org.apache.poi.ss.usermodel.Cell ~> Result[a, String]): Result[a, String] & Impure = 
        use Result.{flatMap};
        let* cell = SheetIO/Internal/Row.getCell(row, ix);
        f(cell)
    

    pub def getRowNum(): RowEvaluator[Int32] = 
        RowEvaluator(row -> Ok(SheetIO/Internal/Row.getRowNum(row)))

    pub def isNullAt(ix: Int32): RowEvaluator[Bool] = 
        liftPrim(x -> withCell(x, ix, ans -> Ok(Object.isNull(ans) as & Impure)))
            `swapError` "isNullAt(${Int32.toString(ix)})"


    pub def getNullAt(ix: Int32): RowEvaluator[Unit] =
        let errMsg = "getNullAt(${Int32.toString(ix)})";
        liftPrim(x -> withCell(x, ix, ans -> if (Object.isNull(ans)) (Ok() as & Impure) else Err(errMsg)))


    pub def getCellTypeAt(ix: Int32): RowEvaluator[CellType] =
        map(SheetIO/CellType.unmarshal, liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getCellType)))
            `swapError` "getCellTypeAt(${Int32.toString(ix)})"

    pub def isBlankAt(ix: Int32): RowEvaluator[Bool] = 
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => return(true)
            case _ => return(false)
        } `swapError` "isBlankAt(${Int32.toString(ix)})"


    pub def getBlankAt(ix: Int32): RowEvaluator[Unit] = 
        let errMsg = "getBlankAt(${Int32.toString(ix)})";
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => return()
            case _ => throwError(errMsg + " - not blank")
        } `swapError` errMsg

    /// Error if the cell type is numeric.
    pub def getStringAt(ix: Int32): RowEvaluator[String] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getStringCellValue))
            `swapError` "getStringAt(${Int32.toString(ix)})"


    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowEvaluator[String] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getDefaultFormatterCellValue))
            `swapError` "getContentAt(${Int32.toString(ix)})"

    pub def getSatifiesAt(ix: Int32, test: String -> Bool): RowEvaluator[String] =
        let step = getContentAt(ix) >>= x -> 
                    if (test(x)) return(x) else throwError("getSatifiesAt('${ix}') - failed");
        step `swapError` "getSatifiesAt('${ix}')"

    pub def getNonEmptyContentAt(ix: Int32): RowEvaluator[String] =
        getSatifiesAt(ix, x -> not String.isEmpty(x)) `swapError` "getNonEmptyContentAt('${ix}')"


    pub def getStringAtWithDefault(ix: Int32, d: String): RowEvaluator[String] =
        let* x = isNullAt(ix);
        let* y = isBlankAt(ix);
        if (x or y) return(d) else getStringAt(ix)

    pub def getStringAtWithDataFormatter(ix: Int32, dfmt: DataFormatter): RowEvaluator[String] =
        let DataFormatter(dfmt1) = dfmt;
        liftPrim(x -> withCell(x, ix, cell -> SheetIO/Internal/Cell.getFormatterCellValue(cell, dfmt1)))


    pub def getBoolAt(ix: Int32): RowEvaluator[Bool] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getBooleanCellValue))
    
    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumericAt(ix: Int32): RowEvaluator[Float64] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getNumericCellValue))


    def getNumAt(ix: Int32, conv: Float64 -> Option[a], errMsg: String): RowEvaluator[a] =
        let* n = getNumericAt(ix);
        match conv(n) {
            case Some(i) => return(i)
            case None => throwError(errMsg)
        }


    pub def getInt8At(ix: Int32): RowEvaluator[Int8] =
        getNumAt(ix, Float64.tryToInt8, "getInt8At")

    pub def getInt16At(ix: Int32): RowEvaluator[Int16] =
        getNumAt(ix, Float64.tryToInt16, "getInt16At")

    pub def getInt32At(ix: Int32): RowEvaluator[Int32] =
        getNumAt(ix, Float64.tryToInt32, "getInt32At")

    pub def getInt64At(ix: Int32): RowEvaluator[Int64] =
        getNumAt(ix, Float64.tryToInt64, "getInt64At")

    pub def getFloat32At(ix: Int32): RowEvaluator[Float32] =
        getNumAt(ix, Float64.tryToFloat32, "getFloat32At")
    
    pub def getFloat64At(ix: Int32): RowEvaluator[Float64] = 
        getNumAt(ix, x -> Some(x), "getFloat64At")

    pub def getLocalDateTimeAt(ix: Int32): RowEvaluator[Time.LocalDateTime] = 
        use Time.LocalDateTime.{LocalDateTime};
        let* dt1 = liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getLocalDateTimeCellValue));
        return(LocalDateTime(dt1))


    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithFormat(ix: Int32, fmt: Time/Format.DateTimeFormatter): RowEvaluator[Time.LocalDateTime] = 
        let* s = getStringAt(ix);
        liftOption(Time/LocalDateTime.parse(s, fmt), "getLocalDateTimeAtWithFormat")
        

    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithPattern(ix: Int32, patt: String): RowEvaluator[Time.LocalDateTime] = 
        let* fmt    = liftOption(Time/Format/DateTimeFormatter.fromPattern(patt), "getLocalDateTimeAtWithPattern - bad pattern");
        let* s      = getStringAt(ix);
        liftOption(Time/LocalDateTime.parse(s, fmt), "getLocalDateTimeAtWithPattern")

    pub def getLocalDateAt(ix: Int32): RowEvaluator[Time.LocalDate] = 
        map(Time/LocalDateTime.toLocalDate, getLocalDateTimeAt(ix))


    /// Cell must be readable as a String
    pub def getLocalDateAtWithFormat(ix: Int32, fmt: Time/Format.DateFormatter): RowEvaluator[Time.LocalDate] = 
        let* s = getStringAt(ix);
        liftOption(Time/LocalDate.parse(s, fmt), "getLocalDateAtWithFormat")
        

    /// Cell must be readable as a String
    pub def getLocalDateAtWithPattern(ix: Int32, patt: String): RowEvaluator[Time.LocalDate] = 
        let* fmt    = liftOption(Time/Format/DateFormatter.fromPattern(patt), "getLocalDateAtWithPattern - bad pattern");
        let* s      = getStringAt(ix);
        liftOption(Time/LocalDate.parse(s, fmt), "getLocalDateAtWithPattern")
        

    def getIfNotEmpty(ix: Int32, getter: Int32 -> RowEvaluator[a]): RowEvaluator[Option[a]] = 
        let* x1 = isNullAt(ix);
        let* y1 = isBlankAt(ix);
        if (x1 or y1) return(None) else map(Some, getter(ix))

    pub def getOptionStringAt(ix: Int32): RowEvaluator[Option[String]] =
        getIfNotEmpty(ix, getStringAt)



    pub def getOptionInt8At(ix: Int32): RowEvaluator[Option[Int8]] =
        getIfNotEmpty(ix, getInt8At)

    pub def getOptionInt16At(ix: Int32): RowEvaluator[Option[Int16]] =
        getIfNotEmpty(ix, getInt16At)

    pub def getOptionInt32At(ix: Int32): RowEvaluator[Option[Int32]] =
        getIfNotEmpty(ix, getInt32At)

    pub def getOptionInt64At(ix: Int32): RowEvaluator[Option[Int64]] =
        getIfNotEmpty(ix, getInt64At)

    pub def getOptionFloat32At(ix: Int32): RowEvaluator[Option[Float32]] =
        getIfNotEmpty(ix, getFloat32At)
    
    pub def getOptionFloat64At(ix: Int32): RowEvaluator[Option[Float64]] = 
        getIfNotEmpty(ix, getFloat64At)

    pub def getOptionLocalDateTimeAt(ix: Int32): RowEvaluator[Option[Time.LocalDateTime]] = 
        getIfNotEmpty(ix, getLocalDateTimeAt)

        
    pub def getFirstCellNum(): RowEvaluator[Int32] =
        liftPrim(r -> SheetIO/Internal/Row.getFirstCellNum(r))

    pub def getLastCellNum(): RowEvaluator[Int32] =
        liftPrim(r -> Result.map(n-> if (n > 0) n - 1 else n,  SheetIO/Internal/Row.getLastCellNum(r)))


    pub def getHeaders(): RowEvaluator[Array[String]] =
        let* end    = getLastCellNum();
        let* xs     = getHeadersHelper(end, 0, ks -> return(ks));
        let* arr    = liftAction(_ -> List.toArray(xs));
        return(arr)
        
    // TODO not tail recursive as it uses let-syntax...
    def getHeadersHelper(end: Int32, ix: Int32, k: List[String] -> RowEvaluator[List[String]]): RowEvaluator[List[String]] =
        if (ix > end)
            k(Nil)
        else 
            let* s = getStringAt(ix);
            getHeadersHelper(end, ix + 1, ss -> k(s :: ss))

    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2]): RowEvaluator[u] =
        ap(map(f, o1), o2)

    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3]): RowEvaluator[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3], 
                  o4: RowEvaluator[t4]): RowEvaluator[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3], 
                  o4: RowEvaluator[t4], 
                  o5: RowEvaluator[t5]): RowEvaluator[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3],
                  o4: RowEvaluator[t4], 
                  o5: RowEvaluator[t5], 
                  o6: RowEvaluator[t6]): RowEvaluator[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3], 
                  o4: RowEvaluator[t4], 
                  o5: RowEvaluator[t5], 
                  o6: RowEvaluator[t6], 
                  o7: RowEvaluator[t7]): RowEvaluator[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3], 
                  o4: RowEvaluator[t4], 
                  o5: RowEvaluator[t5], 
                  o6: RowEvaluator[t6], 
                  o7: RowEvaluator[t7], 
                  o8: RowEvaluator[t8]): RowEvaluator[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef, 
                  o1: RowEvaluator[t1], 
                  o2: RowEvaluator[t2], 
                  o3: RowEvaluator[t3], 
                  o4: RowEvaluator[t4], 
                  o5: RowEvaluator[t5], 
                  o6: RowEvaluator[t6], 
                  o7: RowEvaluator[t7], 
                  o8: RowEvaluator[t8], 
                  o9: RowEvaluator[t9]): RowEvaluator[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef, 
                   o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4], 
                   o5: RowEvaluator[t5], 
                   o6: RowEvaluator[t6], 
                   o7: RowEvaluator[t7], 
                   o8: RowEvaluator[t8], 
                   o9: RowEvaluator[t9], 
                   o10: RowEvaluator[t10]): RowEvaluator[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)


    ///
    /// Returns a 2 element tuple with the results of applying the evaluators `o1` and `o2`.
    ///
    /// Returns an error if evaluating either `o1` or `o2` is an error.
    ///
    pub def tuple2(o1: RowEvaluator[t1], o2: RowEvaluator[t2]): RowEvaluator[(t1, t2)] =
        lift2((x,y) -> (x,y), o1, o2)

    ///
    /// Returns a 3 element tuple with the results of applying the evaluators `o1`, `o2` and `o3`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2` and `o3` is an error.
    ///
    pub def tuple3(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3]): RowEvaluator[(t1, t2, t3)] =
        lift3((x1, x2, x3) -> (x1, x2, x3), o1, o2, o3)

    ///
    /// Returns a 4 element tuple with the results of applying the evaluators `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, `o3` and `o4` is an error.
    ///
    pub def tuple4(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4]): RowEvaluator[(t1, t2, t3, t4)]  =
        lift4((x1, x2, x3, x4) -> (x1, x2, x3, x4), o1, o2, o3, o4)

    ///
    /// Returns a 5 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o5` is an error.
    ///
    pub def tuple5(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4], 
                   o5: RowEvaluator[t5]): RowEvaluator[(t1, t2, t3, t4, t5)]  =
        lift5((x1, x2, x3, x4, x5) -> (x1, x2, x3, x4, x5), o1, o2, o3, o4, o5)

    ///
    /// Returns a 6 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o6` is an error.
    ///
    pub def tuple6(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4], 
                   o5: RowEvaluator[t5], 
                   o6: RowEvaluator[t6]): RowEvaluator[(t1, t2, t3, t4, t5, t6)]  =
        lift6((x1, x2, x3, x4, x5, x6) -> (x1, x2, x3, x4, x5, x6), o1, o2, o3, o4, o5, o6)

    ///
    /// Returns a 7 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o7` is an error.
    ///
    pub def tuple7(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4], 
                   o5: RowEvaluator[t5], 
                   o6: RowEvaluator[t6], 
                   o7: RowEvaluator[t7]): RowEvaluator[(t1, t2, t3, t4, t5, t6, t7)]  =
        lift7((x1, x2, x3, x4, x5, x6, x7) -> (x1, x2, x3, x4, x5, x6, x7), 
                o1, o2, o3, o4, o5, o6, o7)

    ///
    /// Returns a 8 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o8` is an error.
    ///
    pub def tuple8(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4], 
                   o5: RowEvaluator[t5], 
                   o6: RowEvaluator[t6], 
                   o7: RowEvaluator[t7], 
                   o8: RowEvaluator[t8]): RowEvaluator[(t1, t2, t3, t4, t5, t6, t7, t8)]  =
        lift8((x1, x2, x3, x4, x5, x6, x7, x8) -> (x1, x2, x3, x4, x5, x6, x7, x8), 
                o1, o2, o3, o4, o5, o6, o7, o8)

    ///
    /// Returns a 9 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o9` is an error.
    ///
    pub def tuple9(o1: RowEvaluator[t1], 
                   o2: RowEvaluator[t2], 
                   o3: RowEvaluator[t3], 
                   o4: RowEvaluator[t4], 
                   o5: RowEvaluator[t5], 
                   o6: RowEvaluator[t6], 
                   o7: RowEvaluator[t7], 
                   o8: RowEvaluator[t8], 
                   o9: RowEvaluator[t9]): RowEvaluator[(t1, t2, t3, t4, t5, t6, t7, t8, t9)]  =
        lift9((x1, x2, x3, x4, x5, x6, x7, x8, x9) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9)

    ///
    /// Returns a 10 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o10` is an error.
    ///
    pub def tuple10(o1: RowEvaluator[t1], 
                    o2: RowEvaluator[t2], 
                    o3: RowEvaluator[t3], 
                    o4: RowEvaluator[t4], 
                    o5: RowEvaluator[t5], 
                    o6: RowEvaluator[t6], 
                    o7: RowEvaluator[t7], 
                    o8: RowEvaluator[t8], 
                    o9: RowEvaluator[t9], 
                    o10: RowEvaluator[t10]): RowEvaluator[(t1, t2, t3, t4, t5, t6, t7, t8, t9, t10)]  =
        lift10((x1, x2, x3, x4, x5, x6, x7, x8, x9, x10) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9, x10), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9, o10)


    /// Alias for `seqRight`
    pub def *>(p1: RowEvaluator[a], p2: RowEvaluator[b]): RowEvaluator[b] = seqRight(p1, p2)

    /// Alias for `seqLeft`
    pub def <*(p1: RowEvaluator[a], p2: RowEvaluator[b]): RowEvaluator[a] = seqLeft(p1, p2)

    // Haskell's <$>
    pub def <&>(f: a -> b, p2: RowEvaluator[a]): RowEvaluator[b] =
        map(f, p2)

    pub def <&(x: b, p2: RowEvaluator[a]): RowEvaluator[b] =
        p2 *> return(x)

    /// Alias for `bind`
    pub def >>=(ma: RowEvaluator[a], k: a -> RowEvaluator[b]): RowEvaluator[b] = 
        bind(ma, k)

    /// Alias for `flatMap` (flipped `bind`)
    pub def =<<(k: a -> RowEvaluator[b], ma: RowEvaluator[a]): RowEvaluator[b] = 
        flatMap(k, ma)


    // Haskell's ap <*>
    pub def <*>(mf: RowEvaluator[a -> b], p2: RowEvaluator[a]): RowEvaluator[b] =
        ap(mf, p2)

    /// Alias for `kleisliLeft`
    pub def >=>(f1: a -> RowEvaluator[b], f2: b -> RowEvaluator[c]): a -> RowEvaluator[c] = 
        x -> kleisliLeft(f1, f2, x)

    /// Alias for `kleisliRight`
    pub def <=<(f1: b -> RowEvaluator[c], f2: a -> RowEvaluator[b]): a -> RowEvaluator[c] = 
        x -> kleisliRight(f1, f2, x)


}

