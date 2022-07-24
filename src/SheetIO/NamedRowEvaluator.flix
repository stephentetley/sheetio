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
    use SheetIO.DataFormatter;

    /// Includes Map[columName: String => index: Int32]
    pub enum NamedRowEvaluator[r: Region, a: Type](
        Map[String, Int32] -> DataFormatter[r] -> ##org.apache.poi.ss.usermodel.Row -> Result[a, EvalError] & r
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

    instance MonadZero[NamedRowEvaluator[r]] {
        pub def empty(): NamedRowEvaluator[r, a] = SheetIO/NamedRowEvaluator.throwError("empty")
    }

    instance MonadOr[NamedRowEvaluator[r]] {
        pub def orelse(a: NamedRowEvaluator[r, a], b: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
            SheetIO/NamedRowEvaluator.orelse(a, b)
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
    pub def makeHeaders(names: List[String]): Map[String, Int32] =
        let step = (acc, key) -> { 
            let (col, counts, indices) = acc; 
            match Map.get(key, counts) {
                case None => (col + 1, Map.insert(key, 1, counts), Map.insert(key, col, indices))
                case Some(i) => (col + 1, Map.insert(key, i + 1, counts), Map.insert("${key}.${i}", col, indices))
            }
        };        
        let (_, _, ans) = List.foldLeft(step, (0, Map.empty(), Map.empty()), names);
        ans


    pub def point(x: a): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, _) -> Ok(x) as & r)

    def apply1(x: NamedRowEvaluator[r, a], 
               ixs: Map[String, Int32], 
               fmt: DataFormatter[r],
               row: ##org.apache.poi.ss.usermodel.Row): Result[a, EvalError] & r =
        let NamedRowEvaluator(f) = x;
        f(ixs, fmt, row)

    pub def throwError(msg: String): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((_, _, row) -> 
            Err(SheetIO/Internal/EvalError.create(() as Region[r], row, {source = msg, exnMessage = ""}))
        )


    pub def swapError(msg: String, ma: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(ma, ixs, fmt, row)  {
                case Ok(a)    => Ok(a)
                case Err(err) => Err({source = msg | err})
            }
        )

   
    pub def map(f: a -> b & ef, x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => {
                    let ans = f(a) as & r;
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
                        case Ok(a) => Ok(f(a) as & r)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        ) as & ef
    
    pub def flatMap(k: a -> NamedRowEvaluator[r, b] & ef, x: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, b] & ef =
        NamedRowEvaluator((ixs, fmt, row) ->
            match apply1(x, ixs, fmt, row) {
                case Ok(a) => { let ans = k(a) as & r; apply1(ans, ixs, fmt, row) }
                case Err(e) => Err(e)
            }
        ) as & ef



    pub def orelse(ea: NamedRowEvaluator[r, a], eb: NamedRowEvaluator[r, a]): NamedRowEvaluator[r, a] =
        NamedRowEvaluator((ixs, fmt, row) -> 
            match apply1(ea, ixs, fmt, row) {
                case Ok(a)  => Ok(a)
                case Err(_) => apply1(eb, ixs, fmt, row)
            }
        )
   
    def liftPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.DataFormatter -> ##org.apache.poi.ss.usermodel.Row -> Result[a, PrimError] & ef): NamedRowEvaluator[r, a] & ef =
        NamedRowEvaluator((_, fmt, row) -> 
            let r = () as Region[r];
            let DataFormatter(fmt1) = fmt;
            {fn(r, fmt1, row) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.create(r, row))
        )  as & ef

     
    pub def liftAction(fn: Unit -> a & ef): NamedRowEvaluator[r, a] & ef =
        NamedRowEvaluator((_, _, _)  -> 
            let x = fn() as & r; 
            Ok(x)
        ) as & ef


    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): NamedRowEvaluator[r, a] & ef =
        NamedRowEvaluator((_, _, row) -> 
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
        NamedRowEvaluator((ixs, _, row) -> match Map.get(label, ixs) {
            case Some(ix) => Ok(ix)
            case None     => Err(SheetIO/Internal/EvalError.create(() as Region[r], row, {source = "getIndex(${label})", exnMessage = ""}))
        })


    def getCellNamed(label: String): NamedRowEvaluator[r, ##org.apache.poi.ss.usermodel.Cell] \ Read(r) =
        let* ix = getIndex(label); 
        liftPrim((r, _, row) -> SheetIO/Internal/Row.getCell(r, ix, row))
        
    def liftPrimCell(label: String, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.DataFormatter -> ##org.apache.poi.ss.usermodel.Cell -> Result[a, PrimError] & ef): NamedRowEvaluator[r, a] \ { Write(r), ef } =
        let* cell = getCellNamed(label);
        liftPrim((r, fmt, _) -> fn(r, fmt, cell))

    pub def getRowNum(): NamedRowEvaluator[r, Int32] \ Read(r) = 
        liftPrim((r, _, row) -> Ok(SheetIO/Internal/Row.getRowNum(r, row)))


    pub def getSheetName(): NamedRowEvaluator[r, String] \ Read(r) = 
        use Result.{flatMap};
        liftPrim((r, _, row) -> {
            let* sheet = SheetIO/Internal/Row.getSheet(r, row);
            SheetIO/Internal/Sheet.getSheetName(r, sheet)
        })

    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[r, String] \ Read(r) =
        liftPrimCell(label, (r, fmt, cell) -> SheetIO/Internal/DataFormatter.formatCellValue(r, cell, fmt))
            |> swapError("getContentNamed(${label}")


    /// Uses the default DataFormatter
    pub def getTrimmedContentNamed(label: String): NamedRowEvaluator[r, String] \ Read(r) =
        getContentNamed(label) 
            |> map(String.trim)
            |> swapError("getTrimmedContentNamed(${label})")
        

    pub def isNullNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) = 
        Object.isNull <$> getCellNamed(label)

    pub def getNullNamed(label: String): NamedRowEvaluator[r, Unit] \ Read(r) = 
        let* cell = getCellNamed(label);
        if (Object.isNull(cell)) point() else throwError("getNullNamed(${label})")


    pub def getCellTypeNamed(label: String): NamedRowEvaluator[r, CellType] \ Read(r) =
        liftPrimCell(label, (r, _, cell) -> SheetIO/Internal/Cell.getCellType(r, cell)) |> map(SheetIO/CellType.fromNative)

    pub def isBlankNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) = 
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point(true)
            case _ => point(false)
        } |> swapError("isBlankNamed(${label})")


    pub def getBlankNamed(label: String): NamedRowEvaluator[r, Unit] \ Read(r) =
        let errMsg = "isBlankNamed(${label})";
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)


    pub def getStringNamed(label: String): NamedRowEvaluator[r, String] \ Read(r) =
        liftPrimCell(label, (r, _, row) -> SheetIO/Internal/Cell.getStringCellValue(r, row))
            |> swapError("getStringNamed(${label})")
            

    pub def getSatifiesNamed(label: String, test: String -> Bool): NamedRowEvaluator[r, String] \ Read(r) =
        let errMsg = "getSatifiesNamed(${label})";
        let step = () -> getContentNamed(label) >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)



    /// Whitespace is counted as empty, answer is trimmed.
    pub def getNonEmptyContentNamed(label: String): NamedRowEvaluator[r, String] \ Read(r) =
        getSatifiesNamed(label, x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getSatifiesNamed(${label})")


    pub def getStringNamedWithDefault(label: String, d: String): NamedRowEvaluator[r, String] \ Read(r) =
        let* x = isNullNamed(label);
        let* y = isBlankNamed(label);
        if (x or y) point(d) else getStringNamed(label)


    pub def getStringNamedWithDataFormatter(label: String, dfmt: DataFormatter[r]): NamedRowEvaluator[r, String] \ Read(r) =
        let DataFormatter(dfmt1) = dfmt;
        liftPrimCell(label, (r, _, cell) -> SheetIO/Internal/DataFormatter.formatCellValue(r, cell, dfmt1))

    pub def getBoolNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) =
        liftPrimCell(label, (r, _, cell) -> SheetIO/Internal/Cell.getBooleanCellValue(r, cell))
    

    pub def getNumericNamed(label: String): NamedRowEvaluator[r, Float64] \ Read(r) =
        liftPrimCell(label, (r, _, cell) -> SheetIO/Internal/Cell.getNumericCellValue(r, cell))

    def getNumNamed(label: String, conv: Float64 -> Option[a]): NamedRowEvaluator[r, a] \ Read(r) =
        let* n = getNumericNamed(label);
        match conv(n) {
            case Some(d) => point(d)
            case None    => throwError("getNumNamed")
        }


    pub def getInt8Named(label: String): NamedRowEvaluator[r, Int8] \ Read(r) =
        getNumNamed(label, Float64.tryToInt8) 
            |> swapError("getInt8Named(${label})")

    pub def getInt16Named(label: String): NamedRowEvaluator[r, Int16] \ Read(r) =
        getNumNamed(label, Float64.tryToInt16) 
            |> swapError("getInt16Named(${label})")

    pub def getInt32Named(label: String): NamedRowEvaluator[r, Int32] \ Read(r) =
        getNumNamed(label, Float64.tryToInt32) 
            |> swapError("getInt32Named(${label})")

    pub def getInt64Named(label: String): NamedRowEvaluator[r, Int64] \ Read(r) =
        getNumNamed(label, Float64.tryToInt64) 
            |> swapError("getInt64Named(${label})")

    pub def getFloat32Named(label: String): NamedRowEvaluator[r, Float32] \ Read(r) =
        getNumNamed(label, Float64.tryToFloat32) 
            |> swapError("getFloat32Named(${label})")
    
    pub def getFloat64Named(label: String): NamedRowEvaluator[r, Float64] \ Read(r) = 
        getNumericNamed(label) 
            |> swapError("getFloat64Named(${label})")


    def isDateFormattedNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) =
        liftPrimCell(label, (r, _, cell) -> SheetIO/Internal/Util.isCellDateFormatted(r, cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    /// Note - will fail if CellType is String.
    pub def getLocalDateTimeNamed(label: String): NamedRowEvaluator[r, Time.LocalDateTime] \ Read(r)= 
        use Time.LocalDateTime.{LocalDateTime};
        let* ctype = getCellTypeNamed(label);
        if (ctype == TyNumeric) {
            let* dtCell = isDateFormattedNamed(label);
            if (dtCell) {
                let* dt = liftPrimCell(label, (r, _, cell) -> SheetIO/Internal/Cell.getLocalDateTimeCellValue(r, cell));
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeAt - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeAt - not a date")
        } else throwError("getLocalDateTimeAt - not a date / not a numeric cell")


    /// Cell must be readable as a String
    pub def getLocalDateTimeNamedWithFormat(label: String, 
                                            fmt: DateTimeFormatter[Time/Format.DateTimeFormat]): NamedRowEvaluator[r, Time.LocalDateTime] \ Read(r) = 
        let* s = getStringNamed(label);
        liftResult(Time/LocalDateTime.parse(fmt, s))
       

    /// Cell must be readable as a String
    pub def getLocalTimeDateNamedWithPattern(label: String, 
                                                patt: String): NamedRowEvaluator[r, Time.LocalDateTime] \ Read(r) = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringNamed(label);
        liftResult(Time/LocalDateTime.parse(fmt, s))


    pub def getLocalDateNamed(label: String): NamedRowEvaluator[r, Time.LocalDate] \ Read(r) = 
        getLocalDateTimeNamed(label)
            |> map(Time/LocalDateTime.toLocalDate)
            |> swapError("getLocalDateNamed(${label})")

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithFormat(label: String, 
                                        fmt: DateTimeFormatter[Time/Format.DateFormat]): NamedRowEvaluator[r, Time.LocalDate] \ Read(r) = 
        let* s = getStringNamed(label);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[r, Time.LocalDate] \ Read(r) = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringNamed(label);
        liftResult(Time/LocalDate.parse(fmt, s))
    


    def getIfNotEmpty(label: String, getter: String -> NamedRowEvaluator[r, a] \ Read(r)): NamedRowEvaluator[r, Option[a]] \ Read(r) = 
        let* x1 = isNullNamed(label);
        let* y1 = isBlankNamed(label);
        if (x1 or y1) point(None) else map(Some, getter(label))

    pub def getOptionStringNamed(label: String): NamedRowEvaluator[r, Option[String]] \ Read(r) =
        getIfNotEmpty(label, getStringNamed)


    pub def getOptionInt8Named(label: String): NamedRowEvaluator[r, Option[Int8]] \ Read(r) =
        getIfNotEmpty(label, getInt8Named)

    pub def getOptionInt16Named(label: String): NamedRowEvaluator[r, Option[Int16]] \ Read(r) =
        getIfNotEmpty(label, getInt16Named)

    pub def getOptionInt32Named(label: String): NamedRowEvaluator[r, Option[Int32]] \ Read(r) =
        getIfNotEmpty(label, getInt32Named)

    pub def getOptionInt64Named(label: String): NamedRowEvaluator[r, Option[Int64]] \ Read(r) =
        getIfNotEmpty(label, getInt64Named)

    pub def getOptionFloat32Named(label: String): NamedRowEvaluator[r, Option[Float32]] \ Read(r) =
        getIfNotEmpty(label, getFloat32Named)
    
    pub def getOptionFloat64Named(label: String): NamedRowEvaluator[r, Option[Float64]] \ Read(r) = 
        getIfNotEmpty(label, getFloat64Named)
            |> swapError("getOptionFloat64Named(${label})")


    pub def getOptionLocalDateTimeNamed(label: String): NamedRowEvaluator[r, Option[Time.LocalDateTime]] \ Read(r) = 
        getIfNotEmpty(label, getLocalDateTimeNamed)




    def queryFontAt(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] & ef): NamedRowEvaluator[r, a] \ { Read(r), ef } = 
        use Result.flatMap;
        liftPrim((r, _, row) -> 
            let* cell   = SheetIO/Internal/Row.getCell(r, ix, row);
            let* style  = SheetIO/Internal/Cell.getCellStyle(r, cell);
            let* idx    = SheetIO/Internal/CellStyle.getFontIndexAsInt(r, style);
            let* sheet  = SheetIO/Internal/Row.getSheet(r, row);
            let* wb     = SheetIO/Internal/Sheet.getWorkbook(r, sheet);
            let* font   = SheetIO/Internal/Workbook.getFontAt(r, idx, wb);
            let* ans    = fn(r, font); 
            Ok(ans)
        )


    pub def getIsBoldNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getBold)

    pub def getIsItalicNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutNamed(label: String): NamedRowEvaluator[r, Bool] \ Read(r) = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)




}

