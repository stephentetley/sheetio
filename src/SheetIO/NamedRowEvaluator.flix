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

    use SheetIO.DataFormatter;


    /// Includes Map[columName: String => index: Int32]
    pub enum NamedRowEvaluator[r: Region, ka: Type, a: Type](
        Graded.EnvEvaluator[ka, SheetIO/NamedRowEvaluator.Env[r], Interop/Apache/Poi/SS/Usermodel.Row[r], a])


    instance Functor[NamedRowEvaluator[r, ka]] {
        pub def map(f: a -> b & ef, o: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, b] & ef = 
            SheetIO/NamedRowEvaluator.map(f, o)
    }

    instance Applicative[NamedRowEvaluator[r, ka]] {
        pub def point(x: a): NamedRowEvaluator[r, ka, a] = 
            SheetIO/NamedRowEvaluator.point(x)

        pub def ap(f: NamedRowEvaluator[r, ka, a -> b & ef], x: NamedRowEvaluator[r, ka, a]) : NamedRowEvaluator[r, ka, b] & ef = 
            SheetIO/NamedRowEvaluator.ap(f, x)
    }

    instance Monad[NamedRowEvaluator[r, ka]] {
        pub def flatMap(f: a -> NamedRowEvaluator[r, ka, b] & ef, o: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, b] & ef =
            SheetIO/NamedRowEvaluator.flatMap(f, o)
    }

    instance MonadZero[NamedRowEvaluator[r, ka]] {
        pub def empty(): NamedRowEvaluator[r, ka, a] = SheetIO/NamedRowEvaluator.empty()
    }

    instance MonadLib.MonadOr[NamedRowEvaluator[r, ka]] {
        pub def orelse(a: NamedRowEvaluator[r, ka, a], b: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, a] =
            SheetIO/NamedRowEvaluator.orelse(a, b)
    }

    instance MonadLib.MonadError[NamedRowEvaluator[r, ka]] {
        pub def throwError(err: String): NamedRowEvaluator[r, ka, a] =
            SheetIO/NamedRowEvaluator.throwError(err)

        pub def mapError(f: String -> String \ ef, m: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, a] \ef =
            SheetIO/NamedRowEvaluator.mapError(f, m)

        pub def catchError(m: NamedRowEvaluator[r, ka, a], hdl: String -> NamedRowEvaluator[r, ka, a] \ ef): NamedRowEvaluator[r, ka, a] \ ef =
            SheetIO/NamedRowEvaluator.catchError(m, hdl)
    }

}

namespace SheetIO/NamedRowEvaluator {

    
    use GradedError.{swapError}
    use GradedFunctor.{<$>}
    use GradedMonad.{>>=}
    
    use Time/Format.DateTimeFormatter

    use SheetIO.NamedRowEvaluator
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator}

    use Interop/Apache/Poi/SS/Usermodel.Cell
    use Interop/Apache/Poi/SS/Usermodel.DataFormatter
    use Interop/Apache/Poi/SS/Usermodel.Row

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}


    pub type alias Env[r: Region] = 
        { columnIndices = Map[String, Int32]
        , formatter = Interop/Apache/Poi/SS/Usermodel.DataFormatter[r]
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



    pub def getNamedRowEvaluator(x: NamedRowEvaluator[r, ka, a]): MonadLib.EnvEvaluator[r, ka, Env[r], Row[r], a] = 
        let NamedRowEvaluator(x1) = x;
        x1
        
    pub def runNamedRowEvaluator(r: Region[r], 
                                    ma: NamedRowEvaluator[r, a, a],
                                    headings: List[String],
                                    st: Row[r]): Result[a, String] \ r =
        let env = 
            { columnIndices = makeHeaders(headings)
            , formatter = Interop/Apache/Poi/SS/Usermodel/DataFormatter.new(r)
            };
        MonadLib/EnvEvaluator.runEnvEvaluator(r, getNamedRowEvaluator(ma), env, st)


    pub def point(x: a): NamedRowEvaluator[r, ka, a] = 
        NamedRowEvaluator(MonadLib/EnvEvaluator.point(x))


    pub def map(f: a -> b \ ef, ma: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, b] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.map(f, getNamedRowEvaluator(ma)))

    pub def ap(mf: NamedRowEvaluator[r, ka, a -> b \ ef], ma: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, b] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.ap(getNamedRowEvaluator(mf), getNamedRowEvaluator(ma)))

    pub def flatMap(f: a -> NamedRowEvaluator[r, ka, b] \ ef, ma: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, b] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.flatMap(f >> getNamedRowEvaluator, getNamedRowEvaluator(ma)))

    pub def throwError(msg: String): NamedRowEvaluator[r, ka, a] = 
        NamedRowEvaluator(MonadLib/EnvEvaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, a] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.mapError(f, getNamedRowEvaluator(ma)))

    pub def catchError(ma: NamedRowEvaluator[r, ka, a], f: String -> NamedRowEvaluator[r, ka, a] \ ef): NamedRowEvaluator[r, ka, a] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.catchError(getNamedRowEvaluator(ma), f >> getNamedRowEvaluator))

    pub def empty(): NamedRowEvaluator[r, ka, a] = 
        NamedRowEvaluator(MonadLib/EnvEvaluator.empty())
        
    pub def orelse(ma: NamedRowEvaluator[r, ka, a], mb: NamedRowEvaluator[r, ka, a]): NamedRowEvaluator[r, ka, a] =
        NamedRowEvaluator(MonadLib/EnvEvaluator.orelse(getNamedRowEvaluator(ma), getNamedRowEvaluator(mb)))

    def liftGetter(f: Env[r] -> Row[r] -> a \ ef): NamedRowEvaluator[r, ka, a] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.liftGetter(f))
    
    def liftGetterResult(f: Env[r] -> Row[r] -> Result[a, err] \ ef): NamedRowEvaluator[r, ka, a] \ ef with ToString[err]=
        NamedRowEvaluator(MonadLib/EnvEvaluator.liftGetterResult(f))

    pub def liftAction(f: Unit -> a \ ef): NamedRowEvaluator[r, ka, a] \ ef =
        NamedRowEvaluator(MonadLib/EnvEvaluator.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): NamedRowEvaluator[r, ka, a] \ ef with ToString[err] =
        NamedRowEvaluator(MonadLib/EnvEvaluator.liftActionResult(f))
   
     

    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[r, ka, Int32] =
        liftGetterResult((env, _) -> match Map.get(label, env.columnIndices) {
            case Some(ix) => Ok(ix)
            case None     => Err("getIndex(${label}) no index")
        })


    def getCellNamed(label: String): NamedRowEvaluator[r, ka, Cell[r]] \ Read(r) =
        let* ix = getIndex(label); 
        liftGetterResult((_, row) -> Interop/Apache/Poi/SS/Usermodel/Row.getCell(ix, row) |> Option.toOk("getCellNamed(${label})"))


    def liftCellGetter(label: String, 
                        fn: DataFormatter[r] -> Cell[r] -> a \ ef): NamedRowEvaluator[r, ka, a] \ { Read(r), ef } =
        let* ix = getIndex(label);
        liftGetterResult((env, row) -> 
            match Interop/Apache/Poi/SS/Usermodel/Row.getCell(ix, row) {
                case Some(cell) => fn(env.formatter, cell) |> Ok
                case None       => Err("liftCellGetter(${label})")
            }
            
        )

    def liftCellGetterResult(label: String, 
                            fn: DataFormatter[r] -> Cell[r] -> Result[a, err] \ ef): NamedRowEvaluator[r, ka, a] \ { Read(r), ef } with ToString[err] =
        let* ix = getIndex(label);
        liftGetterResult((env, row) -> 
            match Interop/Apache/Poi/SS/Usermodel/Row.getCell(ix, row) {
                case Some(cell) => fn(env.formatter, cell) |> Result.mapErr(ToString.toString)
                case None       => Err("liftCellGetterResult(${label})")
            }
            
        )


    pub def getRowNum(): NamedRowEvaluator[r, ka, Int32] \ Read(r) = 
        liftGetter((_, row) -> Interop/Apache/Poi/SS/Usermodel/Row.getRowNum(row))

    pub def getSheetName(): NamedRowEvaluator[r, ka, String] \ Read(r) = 
        liftGetter((_, row) -> {
            let sheet = Interop/Apache/Poi/SS/Usermodel/Row.getSheet(row);
            Interop/Apache/Poi/SS/Usermodel/Sheet.getSheetName(sheet)
        })


    /// Uses the default DataFormatter
    pub def getContentNamed(label: String): NamedRowEvaluator[r, ka, String] \ Read(r) =
        liftCellGetter(label, (fmt, cell) -> Interop/Apache/Poi/SS/Usermodel/DataFormatter.formatCellValue(cell, fmt))


    /// Uses the default DataFormatter
    pub def getTrimmedContentNamed(label: String): NamedRowEvaluator[r, ka, String] \ Read(r) =
        String.trim <$> getContentNamed(label)
        

    pub def isNullNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) = 
        Object.isNull <$> getCellNamed(label)

    pub def getNullNamed(label: String): NamedRowEvaluator[r, ka, Unit] \ Read(r) = 
        let* cell = getCellNamed(label);
        if (Object.isNull(cell)) point() else throwError("getNullNamed(${label})")


    pub def getCellTypeNamed(label: String): NamedRowEvaluator[r, ka, CellType] \ Read(r) =
        liftCellGetter(label, (_, cell) -> {
            Interop/Apache/Poi/SS/Usermodel/Cell.getCellType(cell) 
                |> SheetIO/CellType.fromNative
        })

    pub def isBlankNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) = 
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point(true)
            case _       => point(false)
        } |> swapError("isBlankNamed(${label})")


    pub def getBlankNamed(label: String): NamedRowEvaluator[r, ka, Unit] \ Read(r) =
        let errMsg = "isBlankNamed(${label})";
        let* typ = getCellTypeNamed(label);
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)


    pub def getStringNamed(label: String): NamedRowEvaluator[r, ka, String] \ Read(r) =
        liftCellGetterResult(label, (_, row) -> Interop/Apache/Poi/SS/Usermodel/Cell.getStringCellValue(row))
            

    pub def getSatifiesNamed(label: String, test: String -> Bool): NamedRowEvaluator[r, ka, String] \ Read(r) =
        let errMsg = "getSatifiesNamed(${label})";
        let step = () -> getContentNamed(label) >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)


    /// Whitespace is counted as empty, answer is trimmed.
    pub def getNonEmptyContentNamed(label: String): NamedRowEvaluator[r, ka, String] \ Read(r) =
        getSatifiesNamed(label, x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getSatifiesNamed(${label})")

    pub def getStringNamedWithDefault(label: String, d: String): NamedRowEvaluator[r, ka, String] \ Read(r) =
        let* x = isNullNamed(label);
        let* y = isBlankNamed(label);
        if (x or y) point(d) else getStringNamed(label)


    pub def getStringNamedWithDataFormatter(label: String, dfmt: DataFormatter[r]): NamedRowEvaluator[r, ka, String] \ Read(r) =
        liftCellGetter(label, (_, cell) -> Interop/Apache/Poi/SS/Usermodel/DataFormatter.formatCellValue(cell, dfmt))

    pub def getBoolNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) =
        liftCellGetterResult(label, (_, cell) -> Interop/Apache/Poi/SS/Usermodel/Cell.getBooleanCellValue(cell))
    

    pub def getNumericNamed(label: String): NamedRowEvaluator[r, ka, Float64] \ Read(r) =
        liftCellGetterResult(label, (_, cell) -> Interop/Apache/Poi/SS/Usermodel/Cell.getNumericCellValue(cell))

    def getNumNamed(label: String, conv: Float64 -> Option[a]): NamedRowEvaluator[r, ka, a] \ Read(r) =
        let* n = getNumericNamed(label);
        match conv(n) {
            case Some(d) => point(d)
            case None    => throwError("getNumNamed")
        }


    pub def getInt8Named(label: String): NamedRowEvaluator[r, ka, Int8] \ Read(r) =
        getNumNamed(label, Float64.tryToInt8) 
            |> swapError("getInt8Named(${label})")

    pub def getInt16Named(label: String): NamedRowEvaluator[r, ka, Int16] \ Read(r) =
        getNumNamed(label, Float64.tryToInt16) 
            |> swapError("getInt16Named(${label})")

    pub def getInt32Named(label: String): NamedRowEvaluator[r, ka, Int32] \ Read(r) =
        getNumNamed(label, Float64.tryToInt32) 
            |> swapError("getInt32Named(${label})")

    pub def getInt64Named(label: String): NamedRowEvaluator[r, ka, Int64] \ Read(r) =
        getNumNamed(label, Float64.tryToInt64) 
            |> swapError("getInt64Named(${label})")

    pub def getFloat32Named(label: String): NamedRowEvaluator[r, ka, Float32] \ Read(r) =
        getNumNamed(label, Float64.tryToFloat32) 
            |> swapError("getFloat32Named(${label})")
    
    pub def getFloat64Named(label: String): NamedRowEvaluator[r, ka, Float64] \ Read(r) = 
        getNumericNamed(label) 
            |> swapError("getFloat64Named(${label})")



    def isDateFormattedNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) =
        liftCellGetter(label, ( _, cell) -> Interop/Apache/Poi/SS/Usermodel/DateUtil.isCellDateFormatted(cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    /// Note - will fail if CellType is String.
    pub def getLocalDateTimeNamed(label: String): NamedRowEvaluator[r, ka, Time.LocalDateTime] \ Read(r)= 
        use Time.LocalDateTime.{LocalDateTime};
        let* ctype = getCellTypeNamed(label);
        if (ctype == TyNumeric) {
            let* dtCell = isDateFormattedNamed(label);
            if (dtCell) {
                let* dt = liftCellGetterResult(label, ( _, cell) -> Interop/Apache/Poi/SS/Usermodel/Cell.getLocalDateTimeCellValue(cell));
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeNamed - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeNamed - not a date")
        } else throwError("getLocalDateTimeNamed - not a date / not a numeric cell")


    /// Cell must be readable as a String
    pub def getLocalDateTimeNamedWithFormat(label: String, 
                                            fmt: DateTimeFormatter[Time/Format.DateTimeFormat]): NamedRowEvaluator[r, ka, Time.LocalDateTime] \ Read(r) = 
        let* s = getStringNamed(label);
        MonadLib/MonadError.liftResult(Time/LocalDateTime.parse(fmt, s))
  
    /// Cell must be readable as a String
    pub def getLocalTimeDateNamedWithPattern(label: String, 
                                                patt: String): NamedRowEvaluator[r, ka, Time.LocalDateTime] \ Read(r) = 
        let* fmt    = MonadLib/MonadError.liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringNamed(label);
        MonadLib/MonadError.liftResult(Time/LocalDateTime.parse(fmt, s))

   
/*  

    pub def getLocalDateNamed(label: String): NamedRowEvaluator[r, ka, Time.LocalDate] \ Read(r) = 
        getLocalDateTimeNamed(label)
            |> map(Time/LocalDateTime.toLocalDate)
            |> swapError("getLocalDateNamed(${label})")

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithFormat(label: String, 
                                        fmt: DateTimeFormatter[Time/Format.DateFormat]): NamedRowEvaluator[r, ka, Time.LocalDate] \ Read(r) = 
        let* s = getStringNamed(label);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateNamedWithPattern(label: String, patt: String): NamedRowEvaluator[r, ka, Time.LocalDate] \ Read(r) = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringNamed(label);
        liftResult(Time/LocalDate.parse(fmt, s))
    


    def getIfNotEmpty(label: String, getter: String -> NamedRowEvaluator[r, ka, a] \ Read(r)): NamedRowEvaluator[r, ka, Option[a]] \ Read(r) = 
        let* x1 = isNullNamed(label);
        let* y1 = isBlankNamed(label);
        if (x1 or y1) point(None) else map(Some, getter(label))

    pub def getOptionStringNamed(label: String): NamedRowEvaluator[r, ka, Option[String]] \ Read(r) =
        getIfNotEmpty(label, getStringNamed)


    pub def getOptionInt8Named(label: String): NamedRowEvaluator[r, ka, Option[Int8]] \ Read(r) =
        getIfNotEmpty(label, getInt8Named)

    pub def getOptionInt16Named(label: String): NamedRowEvaluator[r, ka, Option[Int16]] \ Read(r) =
        getIfNotEmpty(label, getInt16Named)

    pub def getOptionInt32Named(label: String): NamedRowEvaluator[r, ka, Option[Int32]] \ Read(r) =
        getIfNotEmpty(label, getInt32Named)

    pub def getOptionInt64Named(label: String): NamedRowEvaluator[r, ka, Option[Int64]] \ Read(r) =
        getIfNotEmpty(label, getInt64Named)

    pub def getOptionFloat32Named(label: String): NamedRowEvaluator[r, ka, Option[Float32]] \ Read(r) =
        getIfNotEmpty(label, getFloat32Named)
    
    pub def getOptionFloat64Named(label: String): NamedRowEvaluator[r, ka, Option[Float64]] \ Read(r) = 
        getIfNotEmpty(label, getFloat64Named)
            |> swapError("getOptionFloat64Named(${label})")


    pub def getOptionLocalDateTimeNamed(label: String): NamedRowEvaluator[r, ka, Option[Time.LocalDateTime]] \ Read(r) = 
        getIfNotEmpty(label, getLocalDateTimeNamed)




    def queryFontAt(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] & ef): NamedRowEvaluator[r, ka, a] \ { Read(r), ef } = 
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


    pub def getIsBoldNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getBold)

    pub def getIsItalicNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutNamed(label: String): NamedRowEvaluator[r, ka, Bool] \ Read(r) = 
        let* ix = getIndex(label);
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)


*/

}

