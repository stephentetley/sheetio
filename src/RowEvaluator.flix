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

    pub opaque type RowEvaluator[a] = (##org.apache.poi.ss.usermodel.DataFormatter, ##org.apache.poi.ss.usermodel.Row) -> Result[a, String] & Impure


    instance Functor[RowEvaluator] {
        pub def map(f: a -> b & ef, o: RowEvaluator[a]): RowEvaluator[b] & ef = 
            SheetIO/RowEvaluator.map(f, o)
    }

    instance Applicative[RowEvaluator] {
        pub def point(x: a): RowEvaluator[a] = 
            SheetIO/RowEvaluator.point(x)

        pub def ap(f: RowEvaluator[a -> b & ef], x: RowEvaluator[a]) : RowEvaluator[b] & ef = 
            SheetIO/RowEvaluator.ap(f, x)
    }

    instance Monad[RowEvaluator] {
        pub def flatMap(f: a -> RowEvaluator[b] & ef, o: RowEvaluator[a]): RowEvaluator[b] & ef =
            SheetIO/RowEvaluator.flatMap(f, o)
    }

}

namespace SheetIO/RowEvaluator {

    use Monad.{>>=};

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};

    use SheetIO.CellType;
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};
    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};
    
    pub def point(x: a): RowEvaluator[a] =
        RowEvaluator((_, _) -> Ok(x) as & Impure)

    def apply1(x: RowEvaluator[a], 
                fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                row: ##org.apache.poi.ss.usermodel.Row): Result[a, String] & Impure =
        let RowEvaluator(f) = x;
        f(fmt, row)

    pub def throwError(msg: String): RowEvaluator[a] =
        RowEvaluator((_, _) -> Err(msg) as & Impure)

    pub def swapError(ma: RowEvaluator[a], msg: String): RowEvaluator[a] =
        RowEvaluator((fmt, row) ->
            match apply1(ma, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => Err(msg)
            }
        )

    pub def map(f: a -> b & ef, x: RowEvaluator[a]): RowEvaluator[b] & ef =
        RowEvaluator((fmt, row) ->
            match apply1(x, fmt, row) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: RowEvaluator[a -> b & ef], ma: RowEvaluator[a]): RowEvaluator[b] & ef =
        RowEvaluator((fmt, row) ->
            match apply1(mf, fmt, row) {
                case Ok(f) => {
                    match apply1(ma, fmt, row) {
                        case Ok(a) => Ok(f(a) as & Impure)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    

    pub def bind(x: RowEvaluator[a], k: a -> RowEvaluator[b] & ef): RowEvaluator[b] & ef =
        RowEvaluator((fmt, row) ->
            match apply1(x, fmt, row) {
                case Ok(a) => { let ans = k(a) as & Impure; apply1(ans, fmt, row) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(k: a -> RowEvaluator[b] & ef, x: RowEvaluator[a]): RowEvaluator[b] & ef = bind(x, k)


    pub def alt(ea: RowEvaluator[a], eb: RowEvaluator[a]): RowEvaluator[a] =
        RowEvaluator((fmt, row) ->
            match apply1(ea, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => apply1(eb, fmt, row)
            }
        )


    pub def optional(x: RowEvaluator[a]): RowEvaluator[Option[a]] =
        RowEvaluator((fmt, row) -> match apply1(x, fmt, row) {
            case Ok(a) => Ok(Some(a))
            case Err(_) => Ok(None)
        })

    /// Note - this returns the default value for any errors (i.e. unknown column which is a
    /// schema error), not just a "read" error if the cell type is wrong.
    pub def withDefault(x: RowEvaluator[a], d: a): RowEvaluator[a] =
        alt(x, point(d))



    def liftPrim(fn: ##org.apache.poi.ss.usermodel.Row -> Result[a, String] & Impure): RowEvaluator[a] =
        RowEvaluator((_, row) -> fn(row))


    pub def liftAction(fn: Unit -> a & Impure): RowEvaluator[a] =
        RowEvaluator((_, _) -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): RowEvaluator[a] =
        RowEvaluator((_, _) -> fn())


    pub def liftOption(x: Option[a], errMsg: String): RowEvaluator[a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): RowEvaluator[a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }


    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): RowEvaluator[a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
    }

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    def withCell(row: ##org.apache.poi.ss.usermodel.Row,
                 ix: Int32, 
                 f: ##org.apache.poi.ss.usermodel.Cell -> Result[a, String] & Impure): Result[a, String] & Impure = 
        use Result.{flatMap};
        let* cell = SheetIO/Internal/Row.getCell(row, ix);
        f(cell)
    
    def getDataFormatter(): RowEvaluator[##org.apache.poi.ss.usermodel.DataFormatter] = 
        RowEvaluator((fmt, _) -> Ok(fmt) as & Impure)

    pub def getRowNum(): RowEvaluator[Int32] = 
        RowEvaluator((_, row) -> Ok(SheetIO/Internal/Row.getRowNum(row)))

    def decodeColumn(s: String): Result[Int32, String] & Impure = 
        import new org.apache.poi.ss.util.CellAddress(String) as newCellAddress; 
        import org.apache.poi.ss.util.CellAddress.getColumn(); 
        try {
            let addr = newCellAddress(s + "1");
            Ok(getColumn(addr))
        } catch {
            case _: ##java.lang.Exception => Err("decodeColumn") 
        }

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    pub def getColumnIndex(s: String): RowEvaluator[Int32] = 
        liftPrim(_ -> decodeColumn(s))
        


    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowEvaluator[String] =
        let* fmt = getDataFormatter();
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/DataFormatter.formatCellValue(fmt)))
            `swapError` "getContentAt(${Int32.toString(ix)})"

    /// Uses the default DataFormatter
    pub def getTrimmedContentAt(ix: Int32): RowEvaluator[String] =
        map(String.trim, getContentAt(ix)) `swapError` "getTrimmedContentAt(${Int32.toString(ix)})"



    pub def isNullAt(ix: Int32): RowEvaluator[Bool] = 
        liftPrim(x -> withCell(x, ix, ans -> Ok(Object.isNull(ans) as & Impure)))
            `swapError` "isNullAt(${Int32.toString(ix)})"


    pub def getNullAt(ix: Int32): RowEvaluator[Unit] =
        let errMsg = "getNullAt(${Int32.toString(ix)})";
        liftPrim(x -> withCell(x, ix, ans -> if (Object.isNull(ans)) (Ok() as & Impure) else Err(errMsg)))


    pub def getCellTypeAt(ix: Int32): RowEvaluator[CellType] =
        map(SheetIO/CellType.fromNative, liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getCellType)))
            `swapError` "getCellTypeAt(${Int32.toString(ix)})"

    pub def isBlankAt(ix: Int32): RowEvaluator[Bool] = 
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => point(true)
            case _ => point(false)
        } `swapError` "isBlankAt(${Int32.toString(ix)})"


    pub def getBlankAt(ix: Int32): RowEvaluator[Unit] = 
        let errMsg = "getBlankAt(${Int32.toString(ix)})";
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => point()
            case _ => throwError(errMsg + " - not blank")
        } `swapError` errMsg

    /// Error if the cell type is numeric.
    pub def getStringAt(ix: Int32): RowEvaluator[String] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getStringCellValue))
            `swapError` "getStringAt(${Int32.toString(ix)})"


    
    pub def getSatifiesAt(ix: Int32, test: String -> Bool): RowEvaluator[String] =
        let step = getContentAt(ix) >>= x -> 
                    if (test(x)) point(x) else throwError("getSatifiesAt('${ix}') - failed");
        step `swapError` "getSatifiesAt('${ix}')"

    pub def getNonEmptyContentAt(ix: Int32): RowEvaluator[String] =
        (getSatifiesAt(ix, x -> not String.isEmpty(String.trim(x))) `swapError` "getNonEmptyContentAt('${ix}')")
            |> map(String.trim)


    pub def getStringAtWithDefault(ix: Int32, d: String): RowEvaluator[String] =
        let* x = isNullAt(ix);
        let* y = isBlankAt(ix);
        if (x or y) point(d) else getStringAt(ix)

    pub def getStringAtWithDataFormatter(ix: Int32, dfmt: DataFormatter): RowEvaluator[String] =
        let DataFormatter(dfmt1) = dfmt;
        liftPrim(x -> withCell(x, ix, cell -> SheetIO/Internal/DataFormatter.formatCellValue(dfmt1, cell)))


    pub def getBoolAt(ix: Int32): RowEvaluator[Bool] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getBooleanCellValue))
    
    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumericAt(ix: Int32): RowEvaluator[Float64] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getNumericCellValue))


    def getNumAt(ix: Int32, conv: Float64 -> Option[a], errMsg: String): RowEvaluator[a] =
        let* n = getNumericAt(ix);
        match conv(n) {
            case Some(i) => point(i)
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
        let* dt = liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getLocalDateTimeCellValue));
        if (Object.isNull(dt))
            throwError("getLocalDateTimeAt - null")
        else
            point(LocalDateTime(dt))


    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithFormat(ix: Int32, fmt: Time/Format.DateTimeFormatter): RowEvaluator[Time.LocalDateTime] = 
        let* s = getStringAt(ix);
        liftOption(Time/LocalDateTime.parse(fmt, s), "getLocalDateTimeAtWithFormat")
        

    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithPattern(ix: Int32, patt: String): RowEvaluator[Time.LocalDateTime] = 
        let* fmt    = liftOption(Time/Format/DateTimeFormatter.ofPattern(patt), "getLocalDateTimeAtWithPattern - bad pattern");
        let* s      = getStringAt(ix);
        liftOption(Time/LocalDateTime.parse(fmt, s), "getLocalDateTimeAtWithPattern")

    pub def getLocalDateAt(ix: Int32): RowEvaluator[Time.LocalDate] = 
        map(Time/LocalDateTime.toLocalDate, getLocalDateTimeAt(ix))


    /// Cell must be readable as a String
    pub def getLocalDateAtWithFormat(ix: Int32, fmt: Time/Format.DateFormatter): RowEvaluator[Time.LocalDate] = 
        let* s = getStringAt(ix);
        liftOption(Time/LocalDate.parse(fmt, s), "getLocalDateAtWithFormat")
        

    /// Cell must be readable as a String
    pub def getLocalDateAtWithPattern(ix: Int32, patt: String): RowEvaluator[Time.LocalDate] = 
        let* fmt    = liftOption(Time/Format/DateFormatter.ofPattern(patt), "getLocalDateAtWithPattern - bad pattern");
        let* s      = getStringAt(ix);
        liftOption(Time/LocalDate.parse(fmt, s), "getLocalDateAtWithPattern")
        

    def getIfNotEmpty(ix: Int32, getter: Int32 -> RowEvaluator[a]): RowEvaluator[Option[a]] = 
        let* x1 = isNullAt(ix);
        let* y1 = isBlankAt(ix);
        if (x1 or y1) point(None) else map(Some, getter(ix))

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



    def queryFontAt(ix: Int32, asks: ##org.apache.poi.ss.usermodel.Font -> Result[a, String] & Impure): RowEvaluator[a] = 
        use Result.flatMap;
        liftPrim(row -> 
            let* cell   = SheetIO/Internal/Row.getCell(row, ix);
            let* style  = SheetIO/Internal/Cell.getCellStyle(cell);
            let* idx    = SheetIO/Internal/CellStyle.getFontIndexAsInt(style);
            let* sheet  = SheetIO/Internal/Row.getSheet(row);
            let* wb     = SheetIO/Internal/Sheet.getWorkbook(sheet);
            let* font   = SheetIO/Internal/Workbook.getFontAt(wb, idx);
            let* ans    = asks(font); 
            Ok(ans)
        )


    pub def getIsBoldAt(ix: Int32): RowEvaluator[Bool] = 
        queryFontAt(ix, SheetIO/Internal/Font.getBold)


    pub def getIsItalicAt(ix: Int32): RowEvaluator[Bool] = 
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutAt(ix: Int32): RowEvaluator[Bool] = 
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)


    pub def getRowContents(): RowEvaluator[Array[String]] =
        getRowContentsWith(getContentAt)

    pub def getRowContentsWith(evalCell: Int32 -> RowEvaluator[a]): RowEvaluator[Array[a]] =
        let* end    = getLastCellNum();
        let* xs     = getRowContentsWithHelper(evalCell, end, 0, ks -> point(ks));
        let* arr    = liftAction(_ -> List.toArray(xs));
        point(arr)
        
    // Should be tail recursive as bind is in the contiuation
    def getRowContentsWithHelper(evalCell: Int32 -> RowEvaluator[a], 
                                    end: Int32, 
                                    ix: Int32, 
                                    k: List[a] -> RowEvaluator[List[a]]): RowEvaluator[List[a]] =
        if (ix > end)
            k(Nil)
        else 
            getRowContentsWithHelper(evalCell, end, ix + 1, ss -> evalCell(ix) >>= s -> k(s :: ss))

    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(): RowEvaluator[Array[String]] =
        let eval = ix -> alt(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(eval)

    def getIsoLocalDateTimeAsString(ix: Int32): RowEvaluator[String] = 
        let fmt = Time/Format/DateTimeFormatter.iso_local_date_time();
        let* dt = getLocalDateTimeAt(ix);
        let s   = Time/LocalDateTime.format(fmt, dt);
        point(s)


}

