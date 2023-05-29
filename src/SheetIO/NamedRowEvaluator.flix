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




mod SheetIO {

    use SheetIO.DataFormatter


    /// Includes Map[columName: String => index: Int32]
    pub type alias NamedRowEvaluator[ka: Type, ef: Eff, a: Type] = 
        Graded.EnvEvaluator[ka, SheetIO.NamedRowEvaluator.Env, Interop.Apache.Poi.SS.Usermodel.Row, ef, a]


}

mod SheetIO.NamedRowEvaluator {

    
    use GradedFunctor.{map, <$$>}
    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{throwError, swapError}
    use Graded.EnvEvaluator.{liftGetter, liftGetterResult, liftGetterResultWithEnv, liftAction}
    
    use Time.Format.DateTimeFormatter

    use SheetIO.NamedRowEvaluator

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter
    use Interop.Apache.Poi.SS.Usermodel.Row

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    pub type alias Env = 
        { columnIndices = Map[String, Int32]
        , formatter = Interop.Apache.Poi.SS.Usermodel.DataFormatter
        }


    /// Duplicate column names are allowed - subsequent key duplicates are named with a suffixed index:
    /// "name", "name.1", "name.2", , "name.3" ...
    pub def makeHeaders(names: List[String]): Map[String, Int32] =
        let step = (acc, key) -> { 
            let (col, counts, indices) = acc; 
            match Map.get(key, counts) {
                case None    => (col + 1, Map.insert(key, 1, counts), Map.insert(key, col, indices))
                case Some(i) => (col + 1, Map.insert(key, i + 1, counts), Map.insert("${key}.${i}", col, indices))
            }
        };        
        let (_, _, ans) = List.foldLeft(step, (0, Map.empty(), Map.empty()), names);
        ans




    pub def runNamedRowEvaluator(ma: NamedRowEvaluator[a, ef, a],
                                headings: List[String],
                                df: DataFormatter,
                                st: Row): Result[String, a] \ ef =
        let env = 
            { columnIndices = makeHeaders(headings)
            , formatter = df
            };
        Graded.EnvEvaluator.runEnvEvaluator(ma, env, st)


    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[ka, ef, Int32] =
        liftGetterResultWithEnv((env, _) -> checked_ecast(match Map.get(label, env.columnIndices) {
            case Some(ix) => Ok(ix)
            case None     => Err("getIndex(${label}) no index")
        }))


    pub def getCellNamed(label: String): NamedRowEvaluator[ka, ef, Cell] =
        getIndex(label) >>= ix -> 
        liftGetterResult(row -> Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) |> Option.toOk("getCellNamed(${label})"))


    // def liftCellGetter(label: String, 
    //                     fn: DataFormatter -> Cell -> a \ ef): NamedRowEvaluator[ka, ef, a] =
    //     getIndex(label) >>= ix ->
    //     liftGetterResultWithEnv((env, row) -> 
    //         match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
    //             case Some(cell) => fn(env.formatter, cell) |> Ok
    //             case None       => Err("liftCellGetter(${label})")
    //         }
            
    //     )

    // def liftCellGetterResult(label: String, 
    //                         fn: DataFormatter -> Cell -> Result[err, a] \ ef): NamedRowEvaluator[ka, ef, a] with ToString[err] =
    //     getIndex(label) >>= ix ->
    //     liftGetterResultWithEnv((env, row) -> 
    //         match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
    //             case Some(cell) => fn(env.formatter, cell) |> Result.mapErr(ToString.toString)
    //             case None       => Err("liftCellGetterResult(${label})")
    //         }
            
    //     )


    pub def getRowNum(): NamedRowEvaluator[ka, ef, Int32] = 
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row)))

    pub def getSheetName(): NamedRowEvaluator[ka, ef, String] = 
        liftGetter(row -> {
            let sheet = Interop.Apache.Poi.SS.Usermodel.Row.getSheet(row);
            checked_ecast(Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName(sheet))
        })

/*
    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[ka, ef, String] =
        liftCellGetter(label, (fmt, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, fmt))

    /// Uses the default DataFormatter
    pub def getTrimmedContentNamed(label: String): NamedRowEvaluator[r, ka, r, String] =
        getContentNamed(label) <$$> String.trim
        

    pub def isNullNamed(label: String): NamedRowEvaluator[r, ka, r, Bool] = 
        getCellNamed(label) <$$> Object.isNull


    pub def getNullNamed(label: String): NamedRowEvaluator[r, ka, r, Unit] = 
        getCellNamed(label) >>= cell ->
        if (Object.isNull(cell)) point() else throwError("getNullNamed(${label})")


    pub def getCellTypeNamed(label: String): NamedRowEvaluator[r, ka, r, CellType] =
        liftCellGetter(label, (_, cell) -> {
            Interop.Apache.Poi.SS.Usermodel.Cell.getCellType(cell) 
                |> SheetIO.CellType.fromNative
        })

    pub def isBlankNamed(label: String): NamedRowEvaluator[r, ka, r, Bool] = 
        getCellTypeNamed(label) >>= typ ->
        match typ { 
            case TyBlank => point(true)
            case _       => point(false)
        } |> swapError("isBlankNamed(${label})")


    pub def getBlankNamed(label: String): NamedRowEvaluator[r, ka, r, Unit] =
        let errMsg = "isBlankNamed(${label})";
        getCellTypeNamed(label) >>= typ -> 
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)


    pub def getStringNamed(label: String): NamedRowEvaluator[r, ka, r, String] =
        liftCellGetterResult(label, (_, row) -> Interop.Apache.Poi.SS.Usermodel.Cell.getStringCellValue(row))
            

    pub def getSatifiesNamed(label: String, test: String -> Bool): NamedRowEvaluator[r, ka, r, String] =
        let errMsg = "getSatifiesNamed(${label})";
        let step = () -> getContentNamed(label) >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)


    /// Whitespace is counted as empty, answer is trimmed.
    pub def getNonEmptyContentNamed(label: String): NamedRowEvaluator[r, ka, r, String] =
        getSatifiesNamed(label, x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getSatifiesNamed(${label})")

    pub def getStringNamedWithDefault(label: String, d: String): NamedRowEvaluator[r, ka, r, String] =
        isNullNamed(label)  >>= nullc   ->
        isBlankNamed(label) >>= blankc  ->
        if (nullc or blankc) point(d) else getStringNamed(label)


    pub def getStringNamedWithDataFormatter(label: String, dfmt: DataFormatter): NamedRowEvaluator[r, ka, r, String] =
        liftCellGetter(label, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, dfmt))

    pub def getBoolNamed(label: String): NamedRowEvaluator[r, ka, r, Bool] =
        liftCellGetterResult(label, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getBooleanCellValue(cell))
    

    pub def getNumericNamed(label: String): NamedRowEvaluator[r, ka, r, Float64] =
        liftCellGetterResult(label, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getNumericCellValue(cell))

    def getNumNamed(label: String, conv: Float64 -> Option[a]): NamedRowEvaluator[r, ka, r, a] =
        getNumericNamed(label) >>= n ->
        match conv(n) {
            case Some(d) => point(d)
            case None    => throwError("getNumNamed")
        }


    pub def getInt8Named(label: String): NamedRowEvaluator[r, ka, r, Int8] =
        getNumNamed(label, Float64.tryToInt8) 
            |> swapError("getInt8Named(${label})")

    pub def getInt16Named(label: String): NamedRowEvaluator[r, ka, r, Int16] =
        getNumNamed(label, Float64.tryToInt16) 
            |> swapError("getInt16Named(${label})")

    pub def getInt32Named(label: String): NamedRowEvaluator[r, ka, r, Int32] =
        getNumNamed(label, Float64.tryToInt32) 
            |> swapError("getInt32Named(${label})")

    pub def getInt64Named(label: String): NamedRowEvaluator[r, ka, r, Int64] =
        getNumNamed(label, Float64.tryToInt64) 
            |> swapError("getInt64Named(${label})")

    pub def getFloat32Named(label: String): NamedRowEvaluator[r, ka, r, Float32] =
        getNumNamed(label, Float64.tryToFloat32) 
            |> swapError("getFloat32Named(${label})")
    
    pub def getFloat64Named(label: String): NamedRowEvaluator[r, ka, r, Float64] = 
        getNumericNamed(label) 
            |> swapError("getFloat64Named(${label})")



    def isDateFormattedNamed(label: String): NamedRowEvaluator[r, ka, r, Bool] =
        liftCellGetter(label, ( _, cell) -> Interop.Apache.Poi.SS.Usermodel.DateUtil.isCellDateFormatted(cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    /// Note - will fail if CellType is String.
    pub def getLocalDateTimeNamed(label: String): NamedRowEvaluator[r, ka, r, Time.LocalDateTime] = 
        use Time.LocalDateTime.{LocalDateTime};
        getCellTypeNamed(label) >>= ctype ->
        if (ctype == TyNumeric) {
            isDateFormattedNamed(label) >>= dtCell ->
            if (dtCell) {
                liftCellGetterResult(label, ( _, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getLocalDateTimeCellValue(cell)) >>= dt ->
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeNamed - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeNamed - not a date")
        } else throwError("getLocalDateTimeNamed - not a date / not a numeric cell")


    /// Cell must be readable as a String
    pub def getLocalDateTimeNamedWithFormat(label: String, 
                                            fmt: DateTimeFormatter[Time.Format.DateTimeFormat]): NamedRowEvaluator[r, ka, r, Time.LocalDateTime] = 
        getStringNamed(label) >>= s ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))
  
    /// Cell must be readable as a String
    pub def getLocalTimeDateNamedWithPattern(label: String, 
                                                patt: String): NamedRowEvaluator[r, ka, r, Time.LocalDateTime] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getStringNamed(label)                                                   >>= s   -> 
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))

     

    pub def getLocalDateNamed(label: String): NamedRowEvaluator[r, ka, r, Time.LocalDate] = 
        getLocalDateTimeNamed(label)
            |> map(Time.LocalDateTime.toLocalDate)
            |> swapError("getLocalDateNamed(${label})")

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithFormat(label: String, 
                                        fmt: DateTimeFormatter[Time.Format.DateFormat]): NamedRowEvaluator[r, ka, r, Time.LocalDate] = 
        getStringNamed(label) >>= s ->
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[r, ka, r, Time.LocalDate] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt -> 
        getStringNamed(label)                                                   >>= s   -> 
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))
    


    def getIfNotEmpty(label: String, getter: String -> NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + r, Option[a]] = 
        isNullNamed(label)  >>= nullc  ->
        isBlankNamed(label) >>= blankc ->
        if (nullc or blankc) point(None) else map(Some, getter(label))

    pub def getOptionStringNamed(label: String): NamedRowEvaluator[r, ka, r, Option[String]] =
        getIfNotEmpty(label, getStringNamed)


    pub def getOptionInt8Named(label: String): NamedRowEvaluator[r, ka, r, Option[Int8]] =
        getIfNotEmpty(label, getInt8Named)

    pub def getOptionInt16Named(label: String): NamedRowEvaluator[r, ka, r, Option[Int16]] =
        getIfNotEmpty(label, getInt16Named)

    pub def getOptionInt32Named(label: String): NamedRowEvaluator[r, ka, r, Option[Int32]] =
        getIfNotEmpty(label, getInt32Named)

    pub def getOptionInt64Named(label: String): NamedRowEvaluator[r, ka, r, Option[Int64]] =
        getIfNotEmpty(label, getInt64Named)

    pub def getOptionFloat32Named(label: String): NamedRowEvaluator[r, ka, r, Option[Float32]] =
        getIfNotEmpty(label, getFloat32Named)
    
    pub def getOptionFloat64Named(label: String): NamedRowEvaluator[r, ka, r, Option[Float64]] = 
        getIfNotEmpty(label, getFloat64Named)
            |> swapError("getOptionFloat64Named(${label})")


    pub def getOptionLocalDateTimeNamed(label: String): NamedRowEvaluator[r, ka, r, Option[Time.LocalDateTime]] = 
        getIfNotEmpty(label, getLocalDateTimeNamed)




    def queryFontAt(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] \ ef): NamedRowEvaluator[r, ka, a] \ { r, ef } = 
        use Result.flatMap;
        liftPrim((r, _, row) -> 
            let* cell   = SheetIO.Internal.Row.getCell(r, ix, row);
            let* style  = SheetIO.Internal.Cell.getCellStyle(r, cell);
            let* idx    = SheetIO.Internal.CellStyle.getFontIndexAsInt(r, style);
            let* sheet  = SheetIO.Internal.Row.getSheet(r, row);
            let* wb     = SheetIO.Internal.Sheet.getWorkbook(r, sheet);
            let* font   = SheetIO.Internal.Workbook.getFontAt(r, idx, wb);
            let* ans    = fn(r, font); 
            Ok(ans)
        )


    pub def getIsBoldNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ r = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO.Internal.Font.getBold)

    pub def getIsItalicNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ r = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO.Internal.Font.getItalic)

    pub def getIsStrikeoutNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ r = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO.Internal.Font.getStrikeout)


*/

}

