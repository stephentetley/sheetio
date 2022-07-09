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

    use SheetIO/Internal/EvalError.{EvalError};
    
    pub enum RowEvaluator[r: Region, a: Type](
        (##org.apache.poi.ss.usermodel.DataFormatter, ##org.apache.poi.ss.usermodel.Row) -> Result[a, EvalError] & r)


    instance Functor[RowEvaluator[r]] {
        pub def map(f: a -> b & ef, o: RowEvaluator[r, a]): RowEvaluator[r, b] & ef = 
            SheetIO/RowEvaluator.map(f, o)
    }

    instance Applicative[RowEvaluator[r]] {
        pub def point(x: a): RowEvaluator[r, a] = 
            SheetIO/RowEvaluator.point(x)

        pub def ap(f: RowEvaluator[r, a -> b & ef], x: RowEvaluator[r, a]) : RowEvaluator[r, b] & ef = 
            SheetIO/RowEvaluator.ap(f, x)
    }

    instance Monad[RowEvaluator[r]] {
        pub def flatMap(f: a -> RowEvaluator[r, b] & ef, o: RowEvaluator[r, a]): RowEvaluator[r, b] & ef =
            SheetIO/RowEvaluator.flatMap(f, o)
    }

}

namespace SheetIO/RowEvaluator {

    use Monad.{>>=};
    use Functor.{<$>};

    use Time/Format.DateTimeFormatter;

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};

    use SheetIO/Internal/PrimError.{PrimError};
    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO.CellType;
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};
    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};
    
    pub def point(x: a): RowEvaluator[r, a] = RowEvaluator((_, _) -> Ok(x) as & r)

    def apply1(x: RowEvaluator[r, a], 
                fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                row: ##org.apache.poi.ss.usermodel.Row): Result[a, EvalError] & r =
        let RowEvaluator(f) = x;
        f(fmt, row)

/*
    pub def throwError(msg: String): RowEvaluator[r, a] =
        RowEvaluator((_, row) -> Err(SheetIO/EvalError.create(() as Region[r], row, msg)))

    pub def swapError(ma: RowEvaluator[r, a], msg: String): RowEvaluator[r, a] =
        RowEvaluator((fmt, row) ->
            match apply1(ma, fmt, row) {
                case Ok(a)  => Ok(a)
                case Err(_) => Err(SheetIO/EvalError.create(() as Region[r], row, msg))
            }
        )
*/

    pub def map(f: a -> b & ef, x: RowEvaluator[r, a]): RowEvaluator[r, b] & ef =
        RowEvaluator((fmt, row) ->
            match apply1(x, fmt, row) {
                case Ok(a) => {
                    let ans = f(a) as & r;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: RowEvaluator[r, a -> b & ef], ma: RowEvaluator[r, a]): RowEvaluator[r, b] & ef =
        RowEvaluator((fmt, row) ->
            match apply1(mf, fmt, row) {
                case Ok(f) => {
                    match apply1(ma, fmt, row) {
                        case Ok(a) => Ok(f(a) as & r)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    

    pub def bind(x: RowEvaluator[r, a], k: a -> RowEvaluator[r, b] & ef): RowEvaluator[r, b] & ef =
        RowEvaluator((fmt, row) ->
            match apply1(x, fmt, row) {
                case Ok(a) => {let ans = k(a) as & r; apply1(ans, fmt, row)}
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(k: a -> RowEvaluator[r, b] & ef, x: RowEvaluator[r, a]): RowEvaluator[r, b] & ef = bind(x, k)


    pub def alt(ea: RowEvaluator[r, a], eb: RowEvaluator[r, a]): RowEvaluator[r, a] =
        RowEvaluator((fmt, row) ->
            match apply1(ea, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => apply1(eb, fmt, row)
            }
        )


    pub def optional(x: RowEvaluator[r, a]): RowEvaluator[r, Option[a]] =
        RowEvaluator((fmt, row) -> match apply1(x, fmt, row) {
            case Ok(a)  => Ok(Some(a))
            case Err(_) => Ok(None)
        })

    /// Note - this returns the default value for any errors (i.e. unknown column which is a
    /// schema error), not just a "read" error if the cell type is wrong.
    pub def withDefault(x: RowEvaluator[r, a], d: a): RowEvaluator[r, a] =
        alt(x, point(d))


/*

    def liftPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.DataFormatter -> ##org.apache.poi.ss.usermodel.Row -> Result[a, String] & ef): RowEvaluator[r, a] & ef =
        RowEvaluator((fmt, row) -> 
            let r = () as Region[r];
            let ans = fn(r, fmt, row) as & r;
            ans |> Result.mapErr(SheetIO/EvalError.create(r, row))
        )  as & ef


    pub def liftAction(fn: Unit -> a & ef): RowEvaluator[r, a] & ef =
        RowEvaluator((_, _) -> let x = fn() as & r; Ok(x)) as & ef


    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): RowEvaluator[r, a] & ef=
        RowEvaluator((_, row) -> 
            let ans = fn() as & r;
            ans |>  Result.mapErr(SheetIO/EvalError.create(() as Region[r], row))
        ) as & ef


    pub def liftOption(x: Option[a], errMsg: String): RowEvaluator[r, a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): RowEvaluator[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): RowEvaluator[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): RowEvaluator[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    def getCellAt(ix: Int32): RowEvaluator[r, ##org.apache.poi.ss.usermodel.Cell] \ Read(r) = 
        liftPrim((r, _, row) -> SheetIO/Internal/Row.getCell(r, ix, row))

    
    /// def getDataFormatter(): RowEvaluator[r, ##org.apache.poi.ss.usermodel.DataFormatter] = 
    ///     RowEvaluator((fmt, _) -> Ok(fmt) as & Impure)

    pub def getRowNum(): RowEvaluator[r, Int32] \ Read(r) = 
        liftPrim((r, _, row) -> Ok(SheetIO/Internal/Row.getRowNum(r, row)))



    pub def getSheetName(): RowEvaluator[r, String] \ Read(r) = 
        use Result.{flatMap};
        liftPrim((r, _, row) -> {
            let* sheet = SheetIO/Internal/Row.getSheet(r, row);
            SheetIO/Internal/Sheet.getSheetName(r, sheet)
        })
            

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    pub def getColumnIndex(s: String): RowEvaluator[r, Int32] = 
        use Result.{flatMap};
        liftResult(
            region r {            
                let addr = SheetIO/Internal/CellAddress.new(r, s + "1");
                SheetIO/Internal/CellAddress.getColumn(r, addr)
            }
        )
        

    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowEvaluator[r, String] \ Read(r) =
        let* cell = getCellAt(ix);
        liftPrim((r, fmt, _) -> SheetIO/Internal/DataFormatter.formatCellValue(r, cell, fmt))


    /// Uses the default DataFormatter
    pub def getTrimmedContentAt(ix: Int32): RowEvaluator[r, String] \ Read(r) =
        map(String.trim, getContentAt(ix))



    pub def isNullAt(ix: Int32): RowEvaluator[r, Bool] \ Read(r) = 
        Object.isNull <$> getCellAt(ix)



    pub def getNullAt(ix: Int32): RowEvaluator[r, Unit] \ Read(r) = 
        let* cell = getCellAt(ix);
        if (Object.isNull(cell)) point() else throwError("getNullAt(${ix})")



    pub def getCellTypeAt(ix: Int32): RowEvaluator[r, CellType] \ Read(r) =
        let* cell = getCellAt(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.getCellType(r, cell)) |> map(SheetIO/CellType.fromNative)


    pub def isBlankAt(ix: Int32): RowEvaluator[r, Bool] = 
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => point(true)
            case _ => point(false)
        } `swapError` "isBlankAt(${Int32.toString(ix)})"


    pub def getBlankAt(ix: Int32): RowEvaluator[r, Unit] = 
        let errMsg = "getBlankAt(${Int32.toString(ix)})";
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => point()
            case _ => throwError(errMsg + " - not blank")
        } `swapError` errMsg

    /// Error if the cell type is numeric.
    pub def getStringAt(ix: Int32): RowEvaluator[r, String] =
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getStringCellValue))
            `swapError` "getStringAt(${Int32.toString(ix)})"


    
    pub def getSatifiesAt(ix: Int32, test: String -> Bool): RowEvaluator[r, String] =
        let step = getContentAt(ix) >>= x -> 
                    if (test(x)) point(x) else throwError("getSatifiesAt('${ix}') - failed");
        step `swapError` "getSatifiesAt('${ix}')"

    pub def getNonEmptyContentAt(ix: Int32): RowEvaluator[r, String] =
        (getSatifiesAt(ix, x -> not String.isEmpty(String.trim(x))) `swapError` "getNonEmptyContentAt('${ix}')")
            |> map(String.trim)


    pub def getStringAtWithDefault(ix: Int32, d: String): RowEvaluator[r, String] =
        let* x = isNullAt(ix);
        let* y = isBlankAt(ix);
        if (x or y) point(d) else getStringAt(ix)

    pub def getStringAtWithDataFormatter(ix: Int32, dfmt: DataFormatter): RowEvaluator[r, String] =
        let DataFormatter(dfmt1) = dfmt;
        liftPrim(x -> withCell(x, ix, cell -> SheetIO/Internal/DataFormatter.formatCellValue(dfmt1, cell)))



    pub def getBoolAt(ix: Int32): RowEvaluator[r, Bool] \ Read(r) =
        let* cell = getCellAt(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.getBooleanCellValue(r, cell))



    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumericAt(ix: Int32): RowEvaluator[r, Float64] \ Read(r) =
        let* cell = getCellAt(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Cell.getNumericCellValue(r, cell))


    def getNumAt(ix: Int32, conv: Float64 -> Option[a], errMsg: String): RowEvaluator[r, a] \ Read(r) =
        let* n = getNumericAt(ix);
        match conv(n) {
            case Some(i) => point(i)
            case None => throwError(errMsg)
        }


    pub def getInt8At(ix: Int32): RowEvaluator[r, Int8] \ Read(r) =
        getNumAt(ix, Float64.tryToInt8, "getInt8At")

    pub def getInt16At(ix: Int32): RowEvaluator[r, Int16] \ Read(r) =
        getNumAt(ix, Float64.tryToInt16, "getInt16At")

    pub def getInt32At(ix: Int32): RowEvaluator[r, Int32] \ Read(r) =
        getNumAt(ix, Float64.tryToInt32, "getInt32At")

    pub def getInt64At(ix: Int32): RowEvaluator[r, Int64] \ Read(r) =
        getNumAt(ix, Float64.tryToInt64, "getInt64At")

    pub def getFloat32At(ix: Int32): RowEvaluator[r, Float32] \ Read(r) =
        getNumAt(ix, Float64.tryToFloat32, "getFloat32At")
    
    pub def getFloat64At(ix: Int32): RowEvaluator[r, Float64] \ Read(r) = 
        getNumAt(ix, x -> Some(x), "getFloat64At")


    def isDateFormattedAt(ix: Int32): RowEvaluator[r, Bool] \ Read(r) = 
        let* cell = getCellAt(ix);
        liftPrim((r, _, _) -> SheetIO/Internal/Util.isCellDateFormatted(r, cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    ///
    pub def getLocalDateTimeAt(ix: Int32): RowEvaluator[r, Time.LocalDateTime] \ Read(r) = 
        use Time.LocalDateTime.{LocalDateTime};
        let* ctype = getCellTypeAt(ix);
        if (ctype == TyNumeric) {
            let* dtCell = isDateFormattedAt(ix);
            if (dtCell) {
                let* cell = getCellAt(ix);
                let* dt = liftPrim((r, _, _) -> SheetIO/Internal/Cell.getLocalDateTimeCellValue(r, cell));
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeAt - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeAt - not a date")
        } else throwError("getLocalDateTimeAt - not a date / not a numeric cell")



    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithFormat(ix: Int32, fmt: DateTimeFormatter[Time/Format.DateTimeFormat]): RowEvaluator[r, Time.LocalDateTime] = 
        let* s = getStringAt(ix);
        liftResult(Time/LocalDateTime.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithPattern(ix: Int32, patt: String): RowEvaluator[r, Time.LocalDateTime] = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringAt(ix);
        liftResult(Time/LocalDateTime.parse(fmt, s))


    pub def getLocalDateAt(ix: Int32): RowEvaluator[r, Time.LocalDate] = 
        map(Time/LocalDateTime.toLocalDate, getLocalDateTimeAt(ix))



    /// Cell must be readable as a String
    pub def getLocalDateAtWithFormat(ix: Int32, fmt: DateTimeFormatter[Time/Format.DateFormat]): RowEvaluator[r, Time.LocalDate] = 
        let* s = getStringAt(ix);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateAtWithPattern(ix: Int32, patt: String): RowEvaluator[r, Time.LocalDate] = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringAt(ix);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    def getIfNotEmpty(ix: Int32, getter: Int32 -> RowEvaluator[r, a]): RowEvaluator[r, Option[a]] = 
        let* x1 = isNullAt(ix);
        let* y1 = isBlankAt(ix);
        if (x1 or y1) point(None) else map(Some, getter(ix))

    pub def getOptionStringAt(ix: Int32): RowEvaluator[r, Option[String]] =
        getIfNotEmpty(ix, getStringAt)



    pub def getOptionInt8At(ix: Int32): RowEvaluator[r, Option[Int8]] =
        getIfNotEmpty(ix, getInt8At)

    pub def getOptionInt16At(ix: Int32): RowEvaluator[r, Option[Int16]] =
        getIfNotEmpty(ix, getInt16At)

    pub def getOptionInt32At(ix: Int32): RowEvaluator[r, Option[Int32]] =
        getIfNotEmpty(ix, getInt32At)

    pub def getOptionInt64At(ix: Int32): RowEvaluator[r, Option[Int64]] =
        getIfNotEmpty(ix, getInt64At)

    pub def getOptionFloat32At(ix: Int32): RowEvaluator[r, Option[Float32]] =
        getIfNotEmpty(ix, getFloat32At)
    
    pub def getOptionFloat64At(ix: Int32): RowEvaluator[r, Option[Float64]] = 
        getIfNotEmpty(ix, getFloat64At)

    pub def getOptionLocalDateTimeAt(ix: Int32): RowEvaluator[r, Option[Time.LocalDateTime]] = 
        getIfNotEmpty(ix, getLocalDateTimeAt)

        
    pub def getFirstCellNum(): RowEvaluator[r, Int32] =
        liftPrim(r -> SheetIO/Internal/Row.getFirstCellNum(r))

    pub def getLastCellNum(): RowEvaluator[r, Int32] =
        liftPrim(r -> Result.map(n-> if (n > 0) n - 1 else n,  SheetIO/Internal/Row.getLastCellNum(r)))



    def queryFontAt(ix: Int32, asks: ##org.apache.poi.ss.usermodel.Font -> Result[a, String] & Impure): RowEvaluator[r, a] = 
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


    pub def getIsBoldAt(ix: Int32): RowEvaluator[r, Bool] = 
        queryFontAt(ix, SheetIO/Internal/Font.getBold)


    pub def getIsItalicAt(ix: Int32): RowEvaluator[r, Bool] = 
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutAt(ix: Int32): RowEvaluator[r, Bool] = 
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)

    /// Change to List?
    pub def getRowContents(r: Region[r]): RowEvaluator[r, Array[String, r]] \ Write(r) =
        getRowContentsWith(getContentAt, r)

    pub def getRowContentsWith(evalCell: Int32 -> RowEvaluator[r, a], r: Region[r]): RowEvaluator[r, Array[a, r]] \ Write(r) =
        let* end    = getLastCellNum();
        let* xs     = getRowContentsWithHelper(evalCell, end, 0, ks -> point(ks));
        let arr     = List.toArray(r, xs);
        point(arr)
        
    // Should be tail recursive as bind is in the contiuation
    def getRowContentsWithHelper(evalCell: Int32 -> RowEvaluator[r, a], 
                                    end: Int32, 
                                    ix: Int32, 
                                    k: List[a] -> RowEvaluator[r, List[a]]): RowEvaluator[r, List[a]] =
        if (ix > end)
            k(Nil)
        else 
            getRowContentsWithHelper(evalCell, end, ix + 1, ss -> evalCell(ix) >>= s -> k(s :: ss))

    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(r: Region[r]): RowEvaluator[r, Array[String, r]] \ Write(r) =
        let eval = ix -> alt(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(eval, r)

    def getIsoLocalDateTimeAsString(ix: Int32): RowEvaluator[r, String] = 
        let fmt = Time/Format/DateTimeFormatter.iso_local_date_time();
        let* dt = getLocalDateTimeAt(ix);
        let s   = Time/LocalDateTime.format(fmt, dt);
        point(s)


*/
}
