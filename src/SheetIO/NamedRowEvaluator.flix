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



/*
namespace SheetIO {

    use SheetIO/Internal/EvalError.{EvalError};

    /// Includes Map[columName: String => index: Int32]
    pub enum NamedRowEvaluator[r: Region, a: Type](
        (Map[String, Int32], ##org.apache.poi.ss.usermodel.DataFormatter, ##org.apache.poi.ss.usermodel.Row) -> Result[a, EvalError] & r)

    instance Functor[NamedRowEvaluator[r]] {
        pub def map(f: a -> b & ef, o: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef = 
            SheetIO/NamedRowEvaluator.map(f, o)
    }

    instance Applicative[NamedRowEvaluator[r]] {
        pub def point(x: a): NamedRowEvaluator[r, a] = 
            SheetIO/NamedRowEvaluator.point(x)

        pub def ap(f: NamedRowEvaluator[r, a -> b & ef], x: NamedRowEvaluator[r, a]) : NamedRowEvaluator[r, b] & ef = 
            SheetIO/NamedRowEvaluator.ap(f, x)
    }

    instance Monad[NamedRowEvaluator[r]] {
        pub def flatMap(f: a -> NamedRowEvaluator[r, b] & ef, o: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
            SheetIO/NamedRowEvaluator.flatMap(f, o)
    }

}

namespace SheetIO/NamedRowEvaluator {

    use Monad.{>>=};
    
    use Time/Format.DateTimeFormatter;

    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};

    use SheetIO/Internal/EvalError.{EvalError};

    use SheetIO.CellType;
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};
    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};

    /// Duplicate column names are allowed - subsequent key duplicates are named with a suffixed index:
    /// "name", "name.1", "name.2", , "name.3" ...
    pub def makeHeaders(names: Array[String, r]): Map[String, Int32] \ Read(r) =
        let step = (acc, key) -> { 
            let (col, counts, indices) = acc; 
            match Map.get(key, counts) {
                case None => (col + 1, Map.insert(key, 1, counts), Map.insert(key, col, indices))
                case Some(i) => (col + 1, Map.insert(key, i + 1, counts), Map.insert("${key}.${i}", col, indices))
            }
        };        
        let (_, _, ans) = Array.foldLeft(step, (0, Map.empty(), Map.empty()), names);
        ans


    pub def point(x: a): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, _) -> Ok(x) as & Impure)

    def apply1(x: NamedRowEvaluator[r, a], 
               ixs: Map[String, Int32], 
               fmt: ##org.apache.poi.ss.usermodel.DataFormatter,
               row: ##org.apache.poi.ss.usermodel.Row): Result[a, EvalError] & Impure =
        let NamedRowEvaluator(f) = x;
        f(ixs, fmt, row)

    pub def throwError(msg: String): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, row) -> Err(SheetIO/EvalError.create(row, msg)) as & Impure)

    /// TODO - row num is useful but we need a less ad-hoc API...
    pub def throwErrorWithRowNum(msgF: Int32 -> String): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, row) -> 
            let msg = msgF(SheetIO/Internal/Row.getRowNum(row));
            Err(SheetIO/EvalError.create(row, msg))
        )


    pub def swapError(ma: NamedRowEvaluator[r, a], msg: String): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ma, ixs, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => Err(SheetIO/EvalError.create(row, msg))
            }
        )

   pub def mapError(f: String -> String, ma: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ma, ixs, fmt, row) {
                case Ok(a)    => Ok(a)
                case Err(err) => Err({message = f(err.message) | err})
            }
        )


    pub def map(f: a -> b & ef, x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: NamedRowEvaluator[r, a -> b & ef], ma: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(mf, ixs, fmt, row) {
                case Ok(f) => {
                    match apply1(ma, ixs, fmt, row) {
                        case Ok(a) => Ok(f(a) as & Impure)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    

    pub def bind(x: NamedRowEvaluator[r, a], k: a -> NamedRowEvaluator[r, b] & ef): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => { let ans = k(a) as & Impure; apply1(ans, ixs, fmt, row) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(k: a -> NamedRowEvaluator[r, b] & ef, x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef = bind(x, k)


    pub def optional(x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, Option[a]] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => Ok(Some(a))
                case Err(_) => Ok(None)
            }
        )

    /// Note - this returns default for all errors (i.e. unknown column which is a
    /// schema error). Might be better if it can just return default if cell is null...
    pub def withDefault(x: NamedRowEvaluator[r, a], d: a): NamedRowEvaluator[r, a] =
        alt(x, point(d))

    pub def alt(ea: NamedRowEvaluator[r, a], eb: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ea, ixs, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => apply1(eb, ixs, fmt, row) 
            }
        )


    pub def mapM(mf: a -> NamedRowEvaluator[r, b], xs: List[a]): NamedRowEvaluator[r, List[b]] =
        NamedRowEvaluator((ixs, fmt, row) ->
            mapMHelper(ixs, fmt, row, mf, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(ixs: Map[String, Int32], 
                   fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                   row: ##org.apache.poi.ss.usermodel.Row, 
                   mf: a -> NamedRowEvaluator[r, b], 
                   xs: List[a], 
                   fk: EvalError -> Result[List[b], EvalError], 
                   sk: List[b] -> Result[List[b], EvalError]): Result[List[b], EvalError] & Impure =
        match xs {
            case Nil => sk(Nil)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, ixs, fmt, row) {
                    case Err(e) => fk(e)
                    case Ok(x1) => mapMHelper(ixs, fmt, row, mf, rs, fk, ks -> sk(x1 :: ks))
                }
            }
        }
        
        
    def liftPrim(fn: ##org.apache.poi.ss.usermodel.Row -> Result[a, String] & Impure): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, row) -> fn(row) |> Result.mapErr(SheetIO/EvalError.create(row)))


    pub def liftAction(fn: Unit -> a & Impure): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, _) -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, row) -> fn() |> Result.mapErr(SheetIO/EvalError.create(row)))

    pub def liftOption(x: Option[a], errMsg: String): NamedRowEvaluator[r, a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): NamedRowEvaluator[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }
    
    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): NamedRowEvaluator[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): NamedRowEvaluator[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[r, Int32] =
        NamedRowEvaluator((ixs, _, row) -> match Map.get(label, ixs) {
            case Some(ix) => Ok(ix) as & Impure
            case None => Err(SheetIO/EvalError.create(row, "getIndex - '${label}'")) as & Impure
        })

    def getDataFormatter(): NamedRowEvaluator[r, ##org.apache.poi.ss.usermodel.DataFormatter] =
        NamedRowEvaluator((_, fmt, _) -> Ok(fmt) as & Impure)


    def withCell(row: ##org.apache.poi.ss.usermodel.Row,
                 ix: Int32, 
                 f: ##org.apache.poi.ss.usermodel.Cell -> Result[a, String] & Impure): Result[a, String] & Impure = 
        use Result.{flatMap};
        let* cell = SheetIO/Internal/Row.getCell(row, ix);
        f(cell)

    pub def getRowNum(): NamedRowEvaluator[r, Int32] = 
        NamedRowEvaluator((_, _, row) -> Ok(SheetIO/Internal/Row.getRowNum(row)))

    pub def getSheetName(): NamedRowEvaluator[r, String] = 
        NamedRowEvaluator((_, _, row) -> 
            Result.flatMap(SheetIO/Internal/Sheet.getSheetName, SheetIO/Internal/Row.getSheet(row)) 
                |> Result.mapErr(SheetIO/EvalError.create(row)))


    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[r, String] =
        let* ix     = getIndex(label);
        let* fmt    = getDataFormatter();
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/DataFormatter.formatCellValue(fmt))) `swapError` "getContentNamed('${label}')"


    /// Uses the default DataFormatter
    pub def getTrimmedContentNamed(label: String): NamedRowEvaluator[r, String] =
        map(String.trim, getContentNamed(label)) `swapError` "getTrimmedContentNamed(${label})"
        

    pub def isNullNamed(label: String): NamedRowEvaluator[r, Bool] = 
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, ans -> Ok(Object.isNull(ans) as & Impure)))

    pub def getNullNamed(label: String): NamedRowEvaluator[r, Unit] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, ans -> if (Object.isNull(ans)) (Ok() as & Impure) else Err("getNullNamed")))


    pub def getCellTypeNamed(label: String): NamedRowEvaluator[r, CellType] =
        let* ix     = getIndex(label);
        map(SheetIO/CellType.fromNative, liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getCellType)))
            `swapError` "getCellTypeNamed('${label}')"

    pub def isBlankNamed(label: String): NamedRowEvaluator[r, Bool] = 
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point(true)
            case _ => point(false)
        }


    pub def getBlankNamed(label: String): NamedRowEvaluator[r, Unit] = 
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point()
            case _       => throwError("getBlankNamed($label}) - not blank")
        }
    
    pub def getStringNamed(label: String): NamedRowEvaluator[r, String] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getStringCellValue)) `swapError` "getStringNamed('${label}')"
            


    pub def getSatifiesNamed(label: String, test: String -> Bool): NamedRowEvaluator[r, String] =
        let step = getContentNamed(label) >>= x -> 
                    if (test(x)) point(x) else throwError("getSatifiesNamed('${label}') - failed");
        step `swapError` "getSatifiesNamed('${label}')"

    /// Whitespace is counted as empty, answer is trimmed.
    pub def getNonEmptyContentNamed(label: String): NamedRowEvaluator[r, String] =
        (getSatifiesNamed(label, x -> not String.isEmpty(String.trim(x))) `swapError` "getNonEmptyContentNamed('${label}')")
            |> map(String.trim)


    pub def getStringNamedWithDefault(label: String, d: String): NamedRowEvaluator[r, String] =
        let* x = isNullNamed(label);
        if (x) point(d) else getStringNamed(label)


    pub def getStringNamedWithDataFormatter(label: String, dfmt: DataFormatter): NamedRowEvaluator[r, String] =
        let* ix     = getIndex(label);
        let DataFormatter(dfmt1) = dfmt;
        liftPrim(x -> withCell(x, ix, cell -> SheetIO/Internal/DataFormatter.formatCellValue(dfmt1, cell)))


    pub def getBoolNamed(label: String): NamedRowEvaluator[r, Bool] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getBooleanCellValue)) `swapError` "getNumericNamed('${label}')"
    

    pub def getNumericNamed(label: String): NamedRowEvaluator[r, Float64] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getNumericCellValue)) `swapError` "getNumericNamed('${label}')"

    def getNumNamed(label: String, conv: Float64 -> Option[a], errMsg: String): NamedRowEvaluator[r, a] =
        let* n = getNumericNamed(label);
        liftOption(conv(n), errMsg)


    pub def getInt8Named(label: String): NamedRowEvaluator[r, Int8] =
        getNumNamed(label, Float64.tryToInt8, "getInt8Named('${label}')")

    pub def getInt16Named(label: String): NamedRowEvaluator[r, Int16] =
        getNumNamed(label, Float64.tryToInt16, "getInt16Named('${label}')")

    pub def getInt32Named(label: String): NamedRowEvaluator[r, Int32] =
        getNumNamed(label, Float64.tryToInt32, "getInt32Named('${label}')")

    pub def getInt64Named(label: String): NamedRowEvaluator[r, Int64] =
        getNumNamed(label, Float64.tryToInt64, "getInt64Named('${label}')")

    pub def getFloat32Named(label: String): NamedRowEvaluator[r, Float32] =
        getNumNamed(label, Float64.tryToFloat32, "getFloat32Named('${label}')")
    
    pub def getFloat64Named(label: String): NamedRowEvaluator[r, Float64] = 
        getNumNamed(label, x -> Some(x), "getFloat64Named('${label}')")



    def isDateFormattedNamed(label: String): NamedRowEvaluator[r, Bool] = 
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/DateUtil.isCellDateFormatted))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    /// Note - will fail if CellType is String.
    pub def getLocalDateTimeNamed(label: String): NamedRowEvaluator[r, Time.LocalDateTime] = 
        use Time.LocalDateTime.{LocalDateTime};
        let* ctype  = getCellTypeNamed(label);
        let* ix     = getIndex(label);
        if (ctype == TyNumeric) {
            let* dtCell = isDateFormattedNamed(label);
            if (dtCell) {
                let* dt = liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getLocalDateTimeCellValue));
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeAt - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeAt - not a date")
        } else throwError("getLocalDateTimeAt - not a date / not a numeric cell")



    /// Cell must be readable as a String
    pub def getLocalDateTimeNamedWithFormat(label: String, fmt: DateTimeFormatter[Time/Format.DateTimeFormat]): NamedRowEvaluator[r, Time.LocalDateTime] = 
        let* s = getStringNamed(label);
        liftResult(Time/LocalDateTime.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalTimeDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[r, Time.LocalDateTime] = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringNamed(label);
        liftResult(Time/LocalDateTime.parse(fmt, s))


    pub def getLocalDateNamed(label: String): NamedRowEvaluator[r, Time.LocalDate] = 
        mapError(s -> "getLocalDateNamed<<${s}>>",  map(Time/LocalDateTime.toLocalDate, getLocalDateTimeNamed(label)))

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithFormat(label: String, fmt: DateTimeFormatter[Time/Format.DateFormat]): NamedRowEvaluator[r, Time.LocalDate] = 
        let* s = getStringNamed(label);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[r, Time.LocalDate] = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringNamed(label);
        liftResult(Time/LocalDate.parse(fmt, s))
    


    def getIfNotEmpty(label: String, getter: String -> NamedRowEvaluator[r, a]): NamedRowEvaluator[r, Option[a]] = 
        let* x1 = isNullNamed(label);
        let* y1 = isBlankNamed(label);
        if (x1 or y1) point(None) else map(Some, getter(label))

    pub def getOptionStringNamed(label: String): NamedRowEvaluator[r, Option[String]] =
        getIfNotEmpty(label, getStringNamed)



    pub def getOptionInt8Named(label: String): NamedRowEvaluator[r, Option[Int8]] =
        getIfNotEmpty(label, getInt8Named)

    pub def getOptionInt16Named(label: String): NamedRowEvaluator[r, Option[Int16]] =
        getIfNotEmpty(label, getInt16Named)

    pub def getOptionInt32Named(label: String): NamedRowEvaluator[r, Option[Int32]] =
        getIfNotEmpty(label, getInt32Named)

    pub def getOptionInt64Named(label: String): NamedRowEvaluator[r, Option[Int64]] =
        getIfNotEmpty(label, getInt64Named)

    pub def getOptionFloat32Named(label: String): NamedRowEvaluator[r, Option[Float32]] =
        getIfNotEmpty(label, getFloat32Named)
    
    pub def getOptionFloat64Named(label: String): NamedRowEvaluator[r, Option[Float64]] = 
        getIfNotEmpty(label, getFloat64Named) `swapError` ("getOptionFloat64Named('${label}')")

    pub def getOptionLocalDateTimeNamed(label: String): NamedRowEvaluator[r, Option[Time.LocalDateTime]] = 
        getIfNotEmpty(label, getLocalDateTimeNamed)




    def queryFontAt(ix: Int32, asks: ##org.apache.poi.ss.usermodel.Font -> Result[a, String] & Impure): NamedRowEvaluator[r, a] = 
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


    pub def getIsBoldNamed(label: String): NamedRowEvaluator[r, Bool] = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getBold)


    pub def getIsItalicNamed(label: String): NamedRowEvaluator[r, Bool] = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutNamed(label: String): NamedRowEvaluator[r, Bool] = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)


}
*/
