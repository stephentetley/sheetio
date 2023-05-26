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
    pub enum NamedRowEvaluator[r: Region, ka: Type, ef: Eff, a: Type](
        Graded.EnvEvaluator[ka, SheetIO.NamedRowEvaluator.Env, Interop.Apache.Poi.SS.Usermodel.Row[r], ef, a])


    instance GradedFunctor[NamedRowEvaluator[r, ka]] {
        pub def map(f: a -> b \ ef1, 
                    ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1, b] = 
            SheetIO.NamedRowEvaluator.map(f, ma)
    }

    instance GradedApplicative[NamedRowEvaluator[r, ka]] {
        pub def point(x: a): NamedRowEvaluator[r, ka, ef, a] = 
            SheetIO.NamedRowEvaluator.point(x)

        pub def ap(mf: NamedRowEvaluator[r, ka, ef1, a -> b \ ef2], 
                    ma: NamedRowEvaluator[r, ka, ef, a]) : NamedRowEvaluator[r, ka, ef + ef1 + ef2, b] = 
            SheetIO.NamedRowEvaluator.ap(mf, ma)
    }

    instance GradedMonad[NamedRowEvaluator[r, ka]] {
        pub def flatMap(f: a -> NamedRowEvaluator[r, ka, ef1, b] \ ef2, 
                        ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1 + ef2, b] =
            SheetIO.NamedRowEvaluator.flatMap(f, ma)
    }

    instance GradedZero[NamedRowEvaluator[r, ka]] {
        pub def empty(): NamedRowEvaluator[r, ka, ef, a] = 
            SheetIO.NamedRowEvaluator.empty()
    }

    instance GradedOr[NamedRowEvaluator[r, ka]] {
        pub def orelse(ma: NamedRowEvaluator[r, ka, ef1, a], 
                        mb: NamedRowEvaluator[r, ka, ef2, a]): NamedRowEvaluator[r, ka, ef1 + ef2, a] =
            SheetIO.NamedRowEvaluator.orelse(ma, mb)
    }

    instance GradedError[NamedRowEvaluator[r, ka]] {
        pub def throwError(err: String): NamedRowEvaluator[r, ka, ef, a] =
            SheetIO.NamedRowEvaluator.throwError(err)

        pub def mapError(f: String -> String \ ef1, 
                            ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1, a] =
            SheetIO.NamedRowEvaluator.mapError(f, ma)

        pub def catchError(ma: NamedRowEvaluator[r, ka, ef, a], 
                            hdlr: String -> NamedRowEvaluator[r, ka, ef1, a] \ ef2): NamedRowEvaluator[r, ka, ef + ef1 + ef2, a] =
            SheetIO.NamedRowEvaluator.catchError(ma, hdlr)
        
        pub def reifyError(ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef, Result[String, a]] =
            SheetIO.NamedRowEvaluator.reifyError(ma)

    }

}

mod SheetIO.NamedRowEvaluator {

    
    use GradedError.{swapError}
    use GradedFunctor.{<$$>}
    use GradedMonad.{>>=}
    
    use Time.Format.DateTimeFormatter

    use SheetIO.NamedRowEvaluator
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator}

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



    pub def getNamedRowEvaluator(x: NamedRowEvaluator[r, ka, ef, a]): Graded.EnvEvaluator[ka, Env, Row[r], ef, a] = 
        let NamedRowEvaluator(x1) = x;
        x1

    pub def runNamedRowEvaluator(ma: NamedRowEvaluator[r, a, ef, a],
                                headings: List[String],
                                st: Row[r]): Result[String, a] \ ef =
        let env = 
            { columnIndices = makeHeaders(headings)
            , formatter = ???
            };
        Graded.EnvEvaluator.runEnvEvaluator(getNamedRowEvaluator(ma), env, st)


    pub def point(x: a): NamedRowEvaluator[r, ka, ef, a] = 
        NamedRowEvaluator(Graded.EnvEvaluator.point(x))


    pub def map(f: a -> b \ ef1, 
                ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1, b] =
        NamedRowEvaluator(Graded.EnvEvaluator.map(f, getNamedRowEvaluator(ma)))

    pub def ap(mf: NamedRowEvaluator[r, ka, ef1, a -> b \ ef2], 
                ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1 + ef2, b] =
        NamedRowEvaluator(Graded.EnvEvaluator.ap(getNamedRowEvaluator(mf), getNamedRowEvaluator(ma)))

    pub def flatMap(f: a -> NamedRowEvaluator[r, ka, ef1, b] \ ef2, 
                    ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1 + ef2, b] =
        NamedRowEvaluator(Graded.EnvEvaluator.flatMap(f >> getNamedRowEvaluator, getNamedRowEvaluator(ma)))

    pub def throwError(msg: String): NamedRowEvaluator[r, ka, ef, a] = 
        NamedRowEvaluator(Graded.EnvEvaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef1, 
                        ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef + ef1, a] =
        NamedRowEvaluator(Graded.EnvEvaluator.mapError(f, getNamedRowEvaluator(ma)))

    pub def catchError(ma: NamedRowEvaluator[r, ka, ef, a], 
                        f: String -> NamedRowEvaluator[r, ka, ef1, a] \ ef2): NamedRowEvaluator[r, ka, ef + ef1 + ef2, a] =
        NamedRowEvaluator(Graded.EnvEvaluator.catchError(getNamedRowEvaluator(ma), f >> getNamedRowEvaluator))

    pub def reifyError(ma: NamedRowEvaluator[r, ka, ef, a]): NamedRowEvaluator[r, ka, ef, Result[String, a]] =
        NamedRowEvaluator(Graded.EnvEvaluator.reifyError(getNamedRowEvaluator(ma)))


    pub def empty(): NamedRowEvaluator[r, ka, ef, a] = 
        NamedRowEvaluator(Graded.EnvEvaluator.empty())
        
    pub def orelse(ma: NamedRowEvaluator[r, ka, ef1, a], 
                    mb: NamedRowEvaluator[r, ka, ef2, a]): NamedRowEvaluator[r, ka, ef1 + ef2, a] =
        NamedRowEvaluator(Graded.EnvEvaluator.orelse(getNamedRowEvaluator(ma), getNamedRowEvaluator(mb)))

    def liftGetter(f: Row[r] -> a \ ef): NamedRowEvaluator[r, ka, ef, a] =
        NamedRowEvaluator(Graded.EnvEvaluator.liftGetter(f))
    
    def liftGetterResult(f: Row[r] -> Result[err, a] \ ef): NamedRowEvaluator[r, ka, ef, a] with ToString[err] =
        NamedRowEvaluator(Graded.EnvEvaluator.liftGetterResult(f))

    // def _liftGetterWithEnv(f: Env -> Row[r] -> a \ ef): NamedRowEvaluator[r, ka, ef, a] =
    //     NamedRowEvaluator(Graded.EnvEvaluator.liftGetterWithEnv(f))
    
    def liftGetterResultWithEnv(f: Env -> Row[r] -> Result[err, a] \ ef): NamedRowEvaluator[r, ka, ef, a] with ToString[err] =
        NamedRowEvaluator(Graded.EnvEvaluator.liftGetterResultWithEnv(f))


    // pub def _liftAction(f: Unit -> a \ ef): NamedRowEvaluator[r, ka, ef, a] =
    //     NamedRowEvaluator(Graded.EnvEvaluator.liftAction(f))
    
        
    // pub def _liftActionResult(f: Unit -> Result[err, a] \ ef): NamedRowEvaluator[r, ka, ef, a] with ToString[err] =
    //     NamedRowEvaluator(Graded.EnvEvaluator.liftActionResult(f))
   

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[r, ka, Pure, Int32] =
        liftGetterResultWithEnv((env, _) -> match Map.get(label, env.columnIndices) {
            case Some(ix) => Ok(ix)
            case None     => Err("getIndex(${label}) no index")
        })


    def getCellNamed(label: String): NamedRowEvaluator[r, ka, r, Cell[r]] =
        getIndex(label) >>= ix -> 
        liftGetterResult(row -> Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) |> Option.toOk("getCellNamed(${label})"))


    def liftCellGetter(label: String, 
                        fn: DataFormatter -> Cell[r] -> a \ ef): NamedRowEvaluator[r, ka, ef + r, a] =
        getIndex(label) >>= ix ->
        liftGetterResultWithEnv((env, row) -> 
            match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                case Some(cell) => fn(env.formatter, cell) |> Ok
                case None       => Err("liftCellGetter(${label})")
            }
            
        )

    def liftCellGetterResult(label: String, 
                            fn: DataFormatter -> Cell[r] -> Result[err, a] \ ef): NamedRowEvaluator[r, ka, ef + r, a] with ToString[err] =
        getIndex(label) >>= ix ->
        liftGetterResultWithEnv((env, row) -> 
            match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                case Some(cell) => fn(env.formatter, cell) |> Result.mapErr(ToString.toString)
                case None       => Err("liftCellGetterResult(${label})")
            }
            
        )


    pub def getRowNum(): NamedRowEvaluator[r, ka, r, Int32] = 
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row))

    pub def getSheetName(): NamedRowEvaluator[r, ka, r, String] = 
        liftGetter(row -> {
            let sheet = Interop.Apache.Poi.SS.Usermodel.Row.getSheet(row);
            Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName(sheet)
        })


    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[r, ka, r, String] =
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



/*


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

