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

    /// Includes Map[columName: String => index: Int32]
    pub enum NamedRowEvaluator[r: Region, a: Type](
        ##org.apache.poi.ss.usermodel.Row -> ##org.apache.poi.ss.usermodel.DataFormatter -> Map[String, Int32] -> Result[a, EvalError] & r
    )

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

    use Functor.{<$>};
    use Monad.{>>=};
    
    use Time/Format.DateTimeFormatter;

    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};

    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO/Internal/PrimError.{PrimError};

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
        NamedRowEvaluator((_, _, _) -> Ok(x) as & r)

    def apply1(x: NamedRowEvaluator[r, a],
               row: ##org.apache.poi.ss.usermodel.Row, 
               fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
               ixs: Map[String, Int32]): Result[a, EvalError] & r =
        let NamedRowEvaluator(f) = x;
        f(row, fmt, ixs)

    pub def throwError(msg: String): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((row, _, _) -> 
            Err(SheetIO/Internal/EvalError.create(() as Region[r], row, {source = msg, exnMessage = ""}))
        )


    pub def swapError(msg: String, ma: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(ma, row, fmt, ixs)  {
                case Ok(a)    => Ok(a)
                case Err(err) => Err({source = msg | err})
            }
        )

   
    pub def map(f: a -> b & ef, x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(x, row, fmt, ixs) {
                case Ok(a) => {
                    let ans = f(a) as & r;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: NamedRowEvaluator[r, a -> b & ef], ma: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(mf, row, fmt, ixs) {
                case Ok(f) => {
                    match apply1(ma, row, fmt, ixs) {
                        case Ok(a) => Ok(f(a) as & r)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    
    pub def flatMap(k: a -> NamedRowEvaluator[r, b] & ef, x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(x, row, fmt, ixs) {
                case Ok(a) => { let ans = k(a) as & r; apply1(ans, row, fmt, ixs) }
                case Err(e) => Err(e)
            }
        ) as & ef


    pub def bind(x: NamedRowEvaluator[r, a], k: a -> NamedRowEvaluator[r, b] & ef): NamedRowEvaluator[r, b] & ef =
        flatMap(k, x)

    pub def alt(ma: NamedRowEvaluator[r, a], mb: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(ma, row, fmt, ixs) {
                case Ok(a)  => Ok(a)
                case Err(_) => apply1(mb, row, fmt, ixs)
            }
        )


    pub def optional(ma: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, Option[a]] =
        NamedRowEvaluator((row, fmt, ixs) -> match apply1(ma, row, fmt, ixs) {
            case Ok(a)  => Ok(Some(a))
            case Err(_) => Ok(None)
        })

    /// Note - this returns the default value for any errors (i.e. unknown column which is a
    /// schema error), not just a "read" error if the cell type is wrong.
    pub def withDefault(x: NamedRowEvaluator[r, a], d: a): NamedRowEvaluator[r, a] =
        alt(x, point(d))

   
    def liftPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Row -> ##org.apache.poi.ss.usermodel.DataFormatter -> Result[a, PrimError] & ef): NamedRowEvaluator[r, a] & ef =
        NamedRowEvaluator((row, fmt, _) -> 
            let r = () as Region[r];
            {fn(r, row, fmt) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.create(r, row))
        )  as & ef

     
    pub def liftAction(fn: Unit -> a & ef): NamedRowEvaluator[r, a] & ef =
        NamedRowEvaluator((_, _, _)  -> 
            let x = fn() as & r; 
            Ok(x)
        ) as & ef


    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): NamedRowEvaluator[r, a] & ef =
        NamedRowEvaluator((row, _, _) -> 
            let r = () as Region[r]; 
            {fn() as & r} |> Result.mapErr(s -> SheetIO/Internal/EvalError.create(r, row, {source = s, exnMessage = ""}))
        ) as & ef

    pub def liftOption(x: Option[a], errMsg: String): NamedRowEvaluator[r, a] = match x { 
        case Some(x1) => point(x1)
        case None     => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): NamedRowEvaluator[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }
    
    pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): NamedRowEvaluator[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(makeErrMsg: List[e] -> String, x: Validation[a, e]): NamedRowEvaluator[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[r, Int32] =
        NamedRowEvaluator((row, _, ixs) -> match Map.get(label, ixs) {
            case Some(ix) => Ok(ix)
            case None     => Err(SheetIO/Internal/EvalError.create(() as Region[r], row, {source = "getIndex(${label})", exnMessage = ""}))
        })


    def getCellNamed(label: String): NamedRowEvaluator[r, ##org.apache.poi.ss.usermodel.Cell] \ Read(r) =
        let* ix = getIndex(label); 
        liftPrim((r, row, _) -> SheetIO/Internal/Row.getCell(r, ix, row))
        
    def liftPrimCell(label: String, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Cell -> ##org.apache.poi.ss.usermodel.DataFormatter -> Result[a, PrimError] & ef): NamedRowEvaluator[r, a] \ { Write(r), ef } =
        let* cell = getCellNamed(label);
        liftPrim((r, _, fmt) -> fn(r, cell, fmt))

    pub def getRowNum(): NamedRowEvaluator[r, Int32] \ Read(r) = 
        liftPrim((r, row, _) -> Ok(SheetIO/Internal/Row.getRowNum(r, row)))


    pub def getSheetName(): NamedRowEvaluator[r, String] \ Read(r) = 
        use Result.{flatMap};
        liftPrim((r, row, _) -> {
            let* sheet = SheetIO/Internal/Row.getSheet(r, row);
            SheetIO/Internal/Sheet.getSheetName(r, sheet)
        })

    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[r, String] \ Read(r) =
        liftPrimCell(label, SheetIO/Internal/DataFormatter.formatCellValue)
            |> swapError("getContentNamed(${label}")


    /// Uses the default DataFormatter
    pub def getTrimmedContentNamed(label: String): NamedRowEvaluator[r, String] \ Read(r) =
        getContentNamed(label) 
            |> map(String.trim)
            |> swapError("getTrimmedContentNamed(${label})")
        

    pub def isNullNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) = 
        Object.isNull <$> getCellNamed(label)
/*
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


    pub def optional(x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, Option[a]] =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(x, row, fmt, ixs) {
                case Ok(a) => Ok(Some(a))
                case Err(_) => Ok(None)
            }
        )

    /// Note - this returns default for all errors (i.e. unknown column which is a
    /// schema error). Might be better if it can just return default if cell is null...
    pub def withDefault(x: NamedRowEvaluator[r, a], d: a): NamedRowEvaluator[r, a] =
        alt(x, point(d))

    pub def alt(ea: NamedRowEvaluator[r, a], eb: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((row, fmt, ixs) ->
            match apply1(ea, row, fmt, ixs) {
                case Ok(a) => Ok(a)
                case Err(_) => apply1(eb, row, fmt, ixs) 
            }
        )


    pub def mapM(mf: a -> NamedRowEvaluator[r, b], xs: List[a]): NamedRowEvaluator[r, List[b]] =
        NamedRowEvaluator((row, fmt, ixs) ->
            mapMHelper(row, fmt, ixs, mf, xs, e -> Err(e), ys -> Ok(ys))
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
                match apply1(ma, row, fmt, ixs) {
                    case Err(e) => fk(e)
                    case Ok(x1) => mapMHelper(row, fmt, ixs, mf, rs, fk, ks -> sk(x1 :: ks))
                }
            }
        }
        


*/
}

