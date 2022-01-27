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


    pub opaque type NamedRowEvaluator[a] = 
        (Map[String, Int32], ##org.apache.poi.ss.usermodel.DataFormatter, ##org.apache.poi.ss.usermodel.Row) -> Result[a, String] & Impure

    instance Functor[NamedRowEvaluator] {
        pub def map(f: a -> b & ef, o: NamedRowEvaluator[a]): NamedRowEvaluator[b] & ef = 
            SheetIO/NamedRowEvaluator.map(f, o)
    }

    instance Applicative[NamedRowEvaluator] {
        pub def point(x: a): NamedRowEvaluator[a] = 
            SheetIO/NamedRowEvaluator.point(x)

        pub def ap(f: NamedRowEvaluator[a -> b & ef], x: NamedRowEvaluator[a]) : NamedRowEvaluator[b] & ef = 
            SheetIO/NamedRowEvaluator.ap(f, x)
    }

    instance Monad[NamedRowEvaluator] {
        pub def flatMap(f: a -> NamedRowEvaluator[b] & ef, o: NamedRowEvaluator[a]): NamedRowEvaluator[b] & ef =
            SheetIO/NamedRowEvaluator.flatMap(f, o)
    }

}

namespace SheetIO/NamedRowEvaluator {

    use Monad.{>>=};
    
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};

    use SheetIO.CellType;
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};
    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};

    /// Duplicate column names are allowed - subsequent key duplicates are named with a suffixed index:
    /// "name", "name.1", "name.2", , "name.3" ...
    pub def makeHeaders(names: Array[String]): Map[String, Int32] & Impure =
        let step = (acc, key) -> { 
            let (col, counts, indices) = acc; 
            match Map.get(key, counts) {
                case None => (col + 1, Map.insert(key, 1, counts), Map.insert(key, col, indices))
                case Some(i) => (col + 1, Map.insert(key, i + 1, counts), Map.insert("${key}.${i}", col, indices))
            }
        };        
        let (_, _, ans) = Array.foldLeft(step, (0, Map.empty(), Map.empty()), names);
        ans


    pub def point(x: a): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _, _) -> Ok(x) as & Impure)

    def apply1(x: NamedRowEvaluator[a], 
               ixs: Map[String, Int32], 
               fmt: ##org.apache.poi.ss.usermodel.DataFormatter,
               row: ##org.apache.poi.ss.usermodel.Row): Result[a, String] & Impure =
        let NamedRowEvaluator(f) = x;
        f(ixs, fmt, row)

    pub def throwError(msg: String): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _, _) -> Err(msg) as & Impure)

    /// TODO - row num is useful but we need a less ad-hoc API...
    pub def throwErrorWithRowNum(msgF: Int32 -> String): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _, row) -> 
            let msg = msgF(SheetIO/Internal/Row.getRowNum(row));
            Err(msg)
        )


    pub def swapError(ma: NamedRowEvaluator[a], msg: String): NamedRowEvaluator[a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ma, ixs, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => Err(msg)
            }
        )

   pub def mapError(f: String -> String, ma: NamedRowEvaluator[a]): NamedRowEvaluator[a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ma, ixs, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(msg) => Err(f(msg))
            }
        )

    /// TODO - row num is useful but we need a less ad-hoc API...
    pub def swapErrorWithRowNum(ma: NamedRowEvaluator[a], msgF: Int32 -> String): NamedRowEvaluator[a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ma, ixs, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => { 
                    let msg = msgF(SheetIO/Internal/Row.getRowNum(row));
                    Err(msg)
                }
            }
        )


    pub def map(f: a -> b & ef, x: NamedRowEvaluator[a]): NamedRowEvaluator[b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: NamedRowEvaluator[a -> b & ef], ma: NamedRowEvaluator[a]): NamedRowEvaluator[b] & ef =
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
    

    pub def bind(x: NamedRowEvaluator[a], k: a -> NamedRowEvaluator[b] & ef): NamedRowEvaluator[b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => { let ans = k(a) as & Impure; apply1(ans, ixs, fmt, row) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(k: a -> NamedRowEvaluator[b] & ef, x: NamedRowEvaluator[a]): NamedRowEvaluator[b] & ef = bind(x, k)


    pub def optional(x: NamedRowEvaluator[a]): NamedRowEvaluator[Option[a]] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => Ok(Some(a))
                case Err(_) => Ok(None)
            }
        )

    /// Note - this returns default for all errors (i.e. unknown column which is a
    /// schema error). Might be better if it can just return default if cell is null...
    pub def withDefault(x: NamedRowEvaluator[a], d: a): NamedRowEvaluator[a] =
        alt(x, point(d))

    pub def alt(ea: NamedRowEvaluator[a], eb: NamedRowEvaluator[a]): NamedRowEvaluator[a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ea, ixs, fmt, row) {
                case Ok(a) => Ok(a)
                case Err(_) => apply1(eb, ixs, fmt, row) 
            }
        )


    pub def mapM(mf: a -> NamedRowEvaluator[b], xs: List[a]): NamedRowEvaluator[List[b]] =
        NamedRowEvaluator((ixs, fmt, row) ->
            mapMHelper(ixs, fmt, row, mf, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(ixs: Map[String, Int32], 
                   fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                   row: ##org.apache.poi.ss.usermodel.Row, 
                   mf: a -> NamedRowEvaluator[b], 
                   xs: List[a], 
                   fk: String -> Result[List[b], String], 
                   sk: List[b] -> Result[List[b], String]): Result[List[b], String] & Impure =
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
        
        
    def liftPrim(fn: ##org.apache.poi.ss.usermodel.Row -> Result[a, String] & Impure): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _, row) -> fn(row))


    pub def liftAction(fn: Unit -> a & Impure): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _, _) -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _, _) -> fn())

    pub def liftOption(x: Option[a], errMsg: String): NamedRowEvaluator[a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): NamedRowEvaluator[a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): NamedRowEvaluator[a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
    }

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[Int32] =
        NamedRowEvaluator((ixs, _, _) -> match Map.get(label, ixs) {
            case Some(ix) => Ok(ix) as & Impure
            case None => Err("getIndex - '${label}'") as & Impure
        })

    def getDataFormatter(): NamedRowEvaluator[##org.apache.poi.ss.usermodel.DataFormatter] =
        NamedRowEvaluator((_, fmt, _) -> Ok(fmt) as & Impure)


    def withCell(row: ##org.apache.poi.ss.usermodel.Row,
                 ix: Int32, 
                 f: ##org.apache.poi.ss.usermodel.Cell -> Result[a, String] & Impure): Result[a, String] & Impure = 
        use Result.{flatMap};
        let* cell = SheetIO/Internal/Row.getCell(row, ix);
        f(cell)

    pub def getRowNum(): NamedRowEvaluator[Int32] = 
        NamedRowEvaluator((_, _, row) -> Ok(SheetIO/Internal/Row.getRowNum(row)))

    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[String] =
        let* ix     = getIndex(label);
        let* fmt    = getDataFormatter();
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/DataFormatter.formatCellValue(fmt))) `swapError` "getContentNamed('${label}')"


    /// Uses the default DataFormatter
    pub def getTrimmedContentNamed(label: String): NamedRowEvaluator[String] =
        map(String.trim, getContentNamed(label)) `swapError` "getTrimmedContentNamed(${label})"
        

    pub def isNullNamed(label: String): NamedRowEvaluator[Bool] = 
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, ans -> Ok(Object.isNull(ans) as & Impure)))

    pub def getNullNamed(label: String): NamedRowEvaluator[Unit] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, ans -> if (Object.isNull(ans)) (Ok() as & Impure) else Err("getNullNamed")))


    pub def getCellTypeNamed(label: String): NamedRowEvaluator[CellType] =
        let* ix     = getIndex(label);
        map(SheetIO/CellType.fromNative, liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getCellType)))
            `swapError` "getCellTypeNamed('${label}')"

    pub def isBlankNamed(label: String): NamedRowEvaluator[Bool] = 
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point(true)
            case _ => point(false)
        }


    pub def getBlankNamed(label: String): NamedRowEvaluator[Unit] = 
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point()
            case _ => throwError("getBlankNamed($label}) - not blank")
        }
    
    pub def getStringNamed(label: String): NamedRowEvaluator[String] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getStringCellValue)) `swapError` "getStringNamed('${label}')"
            


    pub def getSatifiesNamed(label: String, test: String -> Bool): NamedRowEvaluator[String] =
        let step = getContentNamed(label) >>= x -> 
                    if (test(x)) point(x) else throwError("getSatifiesNamed('${label}') - failed");
        step `swapError` "getSatifiesNamed('${label}')"

    /// Whitespace is counted as empty, answer is trimmed.
    pub def getNonEmptyContentNamed(label: String): NamedRowEvaluator[String] =
        (getSatifiesNamed(label, x -> not String.isEmpty(String.trim(x))) `swapError` "getNonEmptyContentNamed('${label}')")
            |> map(String.trim)


    pub def getStringNamedWithDefault(label: String, d: String): NamedRowEvaluator[String] =
        let* x = isNullNamed(label);
        if (x) point(d) else getStringNamed(label)


    pub def getStringNamedWithDataFormatter(label: String, dfmt: DataFormatter): NamedRowEvaluator[String] =
        let* ix     = getIndex(label);
        let DataFormatter(dfmt1) = dfmt;
        liftPrim(x -> withCell(x, ix, cell -> SheetIO/Internal/DataFormatter.formatCellValue(dfmt1, cell)))


    pub def getBoolNamed(label: String): NamedRowEvaluator[Bool] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getBooleanCellValue)) `swapError` "getNumericNamed('${label}')"
    

    pub def getNumericNamed(label: String): NamedRowEvaluator[Float64] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getNumericCellValue)) `swapError` "getNumericNamed('${label}')"

    def getNumNamed(label: String, conv: Float64 -> Option[a], errMsg: String): NamedRowEvaluator[a] =
        let* n = getNumericNamed(label);
        liftOption(conv(n), errMsg)


    pub def getInt8Named(label: String): NamedRowEvaluator[Int8] =
        getNumNamed(label, Float64.tryToInt8, "getInt8Named('${label}')")

    pub def getInt16Named(label: String): NamedRowEvaluator[Int16] =
        getNumNamed(label, Float64.tryToInt16, "getInt16Named('${label}')")

    pub def getInt32Named(label: String): NamedRowEvaluator[Int32] =
        getNumNamed(label, Float64.tryToInt32, "getInt32Named('${label}')")

    pub def getInt64Named(label: String): NamedRowEvaluator[Int64] =
        getNumNamed(label, Float64.tryToInt64, "getInt64Named('${label}')")

    pub def getFloat32Named(label: String): NamedRowEvaluator[Float32] =
        getNumNamed(label, Float64.tryToFloat32, "getFloat32Named('${label}')")
    
    pub def getFloat64Named(label: String): NamedRowEvaluator[Float64] = 
        getNumNamed(label, x -> Some(x), "getFloat64Named('${label}')")



    def isDateFormattedNamed(label: String): NamedRowEvaluator[Bool] = 
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/DateUtil.isCellDateFormatted))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    /// Note - will fail if CellType is String.
    pub def getLocalDateTimeNamed(label: String): NamedRowEvaluator[Time.LocalDateTime] = 
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
    pub def getLocalDateTimeNamedWithFormat(label: String, fmt: Time/Format.DateTimeFormatter): NamedRowEvaluator[Time.LocalDateTime] = 
        let* s = getStringNamed(label);
        liftOption(Time/LocalDateTime.parse(fmt, s), "getLocalDateTimeNamedWithFormat - '${s}'")
        

    /// Cell must be readable as a String
    pub def getLocalTimeDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[Time.LocalDateTime] = 
        let* fmt    = liftOption(Time/Format/DateTimeFormatter.ofPattern(patt), "getLocalTimeDateNamedWithPattern - bad pattern");
        let* s      = getStringNamed(label);
        liftOption(Time/LocalDateTime.parse(fmt, s), "getLocalTimeDateNamedWithPattern - '${s}'")


    pub def getLocalDateNamed(label: String): NamedRowEvaluator[Time.LocalDate] = 
        mapError(s -> "getLocalDateNamed<<${s}>>",  map(Time/LocalDateTime.toLocalDate, getLocalDateTimeNamed(label)))

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithFormat(label: String, fmt: Time/Format.DateFormatter): NamedRowEvaluator[Time.LocalDate] = 
        let* s = getStringNamed(label);
        liftOption(Time/LocalDate.parse(fmt, s), "getLocalDateNamedWithFormat - '${s}'")
        

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[Time.LocalDate] = 
        let* fmt    = liftOption(Time/Format/DateFormatter.ofPattern(patt), "getLocalDateNamedWithPattern - bad pattern");
        let* s      = getStringNamed(label);
        liftOption(Time/LocalDate.parse(fmt, s), "getLocalDateNamedWithPattern - '${s}'")
    


    def getIfNotEmpty(label: String, getter: String -> NamedRowEvaluator[a]): NamedRowEvaluator[Option[a]] = 
        let* x1 = isNullNamed(label);
        let* y1 = isBlankNamed(label);
        if (x1 or y1) point(None) else map(Some, getter(label))

    pub def getOptionStringNamed(label: String): NamedRowEvaluator[Option[String]] =
        getIfNotEmpty(label, getStringNamed)



    pub def getOptionInt8Named(label: String): NamedRowEvaluator[Option[Int8]] =
        getIfNotEmpty(label, getInt8Named)

    pub def getOptionInt16Named(label: String): NamedRowEvaluator[Option[Int16]] =
        getIfNotEmpty(label, getInt16Named)

    pub def getOptionInt32Named(label: String): NamedRowEvaluator[Option[Int32]] =
        getIfNotEmpty(label, getInt32Named)

    pub def getOptionInt64Named(label: String): NamedRowEvaluator[Option[Int64]] =
        getIfNotEmpty(label, getInt64Named)

    pub def getOptionFloat32Named(label: String): NamedRowEvaluator[Option[Float32]] =
        getIfNotEmpty(label, getFloat32Named)
    
    pub def getOptionFloat64Named(label: String): NamedRowEvaluator[Option[Float64]] = 
        getIfNotEmpty(label, getFloat64Named) `swapErrorWithRowNum` (i -> "getOptionFloat64Named('${label}') - row: ${i}")

    pub def getOptionLocalDateTimeNamed(label: String): NamedRowEvaluator[Option[Time.LocalDateTime]] = 
        getIfNotEmpty(label, getLocalDateTimeNamed)




    def queryFontAt(ix: Int32, asks: ##org.apache.poi.ss.usermodel.Font -> Result[a, String] & Impure): NamedRowEvaluator[a] = 
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


    pub def getIsBoldNamed(label: String): NamedRowEvaluator[Bool] = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getBold)


    pub def getIsItalicNamed(label: String): NamedRowEvaluator[Bool] = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutNamed(label: String): NamedRowEvaluator[Bool] = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)


}

