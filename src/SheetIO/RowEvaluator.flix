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
    

    // pass around DataFormatter[r]
    pub enum RowEvaluator[r: Region, ka: Type, a: Type](
        Graded.EnvEvaluator[r, ka, Interop/Apache/Poi/SS/Usermodel.DataFormatter[r], Interop/Apache/Poi/SS/Usermodel.Row[r], a])

    instance Functor[RowEvaluator[r, ka]] {
        pub def map(f: a -> b & ef, o: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, b] & ef = 
            SheetIO/RowEvaluator.map(f, o)
    }

    instance Applicative[RowEvaluator[r, ka]] {
        pub def point(x: a): RowEvaluator[r, ka, a] = 
            SheetIO/RowEvaluator.point(x)

        pub def ap(f: RowEvaluator[r, ka, a -> b & ef], x: RowEvaluator[r, ka, a]) : RowEvaluator[r, ka, b] & ef = 
            SheetIO/RowEvaluator.ap(f, x)
    }

    instance Monad[RowEvaluator[r, ka]] {
        pub def flatMap(f: a -> RowEvaluator[r, ka, b] & ef, o: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, b] & ef =
            SheetIO/RowEvaluator.flatMap(f, o)
    }

    instance MonadZero[RowEvaluator[r, ka]] {
        pub def empty(): RowEvaluator[r, ka, a] = SheetIO/RowEvaluator.throwError("empty")
    }

    instance Graded.MonadOr[RowEvaluator[r, ka]] {
        pub def orelse(a: RowEvaluator[r, ka, a], b: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, a] =
            SheetIO/RowEvaluator.orelse(a, b)
    }

    instance Graded.MonadError[RowEvaluator[r, ka]] {
        pub def throwError(err: String): RowEvaluator[r, ka, a] =
            SheetIO/RowEvaluator.throwError(err)

        pub def mapError(f: String -> String \ ef, m: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, a] \ef =
            SheetIO/RowEvaluator.mapError(f, m)

        pub def catchError(m: RowEvaluator[r, ka, a], hdl: String -> RowEvaluator[r, ka, a] \ ef): RowEvaluator[r, ka, a] \ ef =
            SheetIO/RowEvaluator.catchError(m, hdl)
    }
}

namespace SheetIO/RowEvaluator {

    
    use GradedError.{swapError}
    use GradedMonad.{>>=}
    use GradedFunctor.{<$>}
    use Time/Format.DateTimeFormatter
    
    use SheetIO.RowEvaluator
    use SheetIO.RowEvaluator.{RowEvaluator}

    use Interop/Apache/Poi/SS/Usermodel.Cell
    use Interop/Apache/Poi/SS/Usermodel.DataFormatter
    use Interop/Apache/Poi/SS/Usermodel.Row

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    


    pub def getRowEvaluator(x: RowEvaluator[r, ka, a]): Graded.EnvEvaluator[r, ka, DataFormatter[r], Row[r], a] = 
        let RowEvaluator(x1) = x;
        x1
        
    pub def runRowEvaluator(r: Region[r], 
                                ma: RowEvaluator[r, a, a], 
                                st: Row[r]): Result[a, String] \ r =
        let env = Interop/Apache/Poi/SS/Usermodel/DataFormatter.new(r);
        Graded/EnvEvaluator.runEnvEvaluator(r, getRowEvaluator(ma), env, st)


    pub def point(x: a): RowEvaluator[r, ka, a] = 
        RowEvaluator(Graded/EnvEvaluator.point(x))

    pub def map(f: a -> b \ ef, ma: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, b] \ ef =
        RowEvaluator(Graded/EnvEvaluator.map(f, getRowEvaluator(ma)))

    pub def ap(mf: RowEvaluator[r, ka, a -> b \ ef], ma: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, b] \ ef =
        RowEvaluator(Graded/EnvEvaluator.ap(getRowEvaluator(mf), getRowEvaluator(ma)))

    pub def flatMap(f: a -> RowEvaluator[r, ka, b] \ ef, ma: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, b] \ ef =
        RowEvaluator(Graded/EnvEvaluator.flatMap(f >> getRowEvaluator, getRowEvaluator(ma)))

    pub def throwError(msg: String): RowEvaluator[r, ka, a] = 
        RowEvaluator(Graded/EnvEvaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, a] \ ef =
        RowEvaluator(Graded/EnvEvaluator.mapError(f, getRowEvaluator(ma)))

    pub def catchError(ma: RowEvaluator[r, ka, a], f: String -> RowEvaluator[r, ka, a] \ ef): RowEvaluator[r, ka, a] \ ef =
        RowEvaluator(Graded/EnvEvaluator.catchError(getRowEvaluator(ma), f >> getRowEvaluator))

    pub def empty(): RowEvaluator[r, ka, a] = 
        RowEvaluator(Graded/EnvEvaluator.empty())
        
    pub def orelse(ma: RowEvaluator[r, ka, a], mb: RowEvaluator[r, ka, a]): RowEvaluator[r, ka, a] =
        RowEvaluator(Graded/EnvEvaluator.orelse(getRowEvaluator(ma), getRowEvaluator(mb)))

    def liftGetter(f: DataFormatter[r] -> Row[r] -> a \ ef): RowEvaluator[r, ka, a] \ ef =
        RowEvaluator(Graded/EnvEvaluator.liftGetter(f))
    
    def liftGetterResult(f: DataFormatter[r] -> Row[r] -> Result[a, err] \ ef): RowEvaluator[r, ka, a] \ ef with ToString[err] =
        RowEvaluator(Graded/EnvEvaluator.liftGetterResult(f))



    pub def liftAction(f: Unit -> a \ ef): RowEvaluator[r, ka, a] \ ef =
        RowEvaluator(Graded/EnvEvaluator.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): RowEvaluator[r, ka, a] \ ef with ToString[err] =
        RowEvaluator(Graded/EnvEvaluator.liftActionResult(f))



    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    def getCellAt(ix: Int32): RowEvaluator[r, ka, Cell[r]] \ Read(r) = 
        liftGetterResult((_, row) -> Interop/Apache/Poi/SS/Usermodel/Row.getCell(ix, row) |> Option.toOk("getCellAt(${ix})"))


    def liftCellGetter(ix: Int32, 
                        fn: DataFormatter[r] -> Cell[r] -> a \ ef): RowEvaluator[r, ka, a] \ { Read(r), ef } =
        liftGetterResult((fmt, row) -> 
            match Interop/Apache/Poi/SS/Usermodel/Row.getCell(ix, row) {
                case Some(cell) => fn(fmt, cell) |> Ok
                case None       => Err("liftCellGetter(${ix})")
            }
            
        )

    def liftCellGetterResult(ix: Int32, 
                            fn: DataFormatter[r] -> Cell[r] -> Result[a, err] \ ef): RowEvaluator[r, ka, a] \ { Read(r), ef } with ToString[err] =
        liftGetterResult((fmt, row) -> 
            match Interop/Apache/Poi/SS/Usermodel/Row.getCell(ix, row) {
                case Some(cell) => fn(fmt, cell) |> Result.mapErr(ToString.toString)
                case None       => Err("liftCellGetterResult(${ix})")
            }
            
        )

    pub def getRowNum(): RowEvaluator[r, ka, Int32] \ Read(r) = 
        liftGetter((_, row) -> Interop/Apache/Poi/SS/Usermodel/Row.getRowNum(row))


    pub def getSheetName(): RowEvaluator[r, ka, String] \ Read(r) = 
        liftGetter((_, row) -> {
            let sheet = Interop/Apache/Poi/SS/Usermodel/Row.getSheet(row);
            Interop/Apache/Poi/SS/Usermodel/Sheet.getSheetName(sheet)
        })
            

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    pub def getColumnIndex(s: String): RowEvaluator[r, ka, Int32] = 
        liftAction(_ -> 
            let addr = Interop/Apache/Poi/SS/Util/CellAddress.cellAddressWithAddress(s + "1");
            Interop/Apache/Poi/SS/Util/CellAddress.getColumn(addr)
        )


    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowEvaluator[r, ka, String] \ Read(r) =
        liftCellGetter(ix, (fmt, cell) -> Interop/Apache/Poi/SS/Usermodel/DataFormatter.formatCellValue(cell, fmt))


    /// Uses the default DataFormatter
    pub def getTrimmedContentAt(ix: Int32): RowEvaluator[r, ka, String] \ Read(r) =
        map(String.trim, getContentAt(ix))

    pub def isNullAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        Object.isNull <$> getCellAt(ix)


    pub def getNullAt(ix: Int32): RowEvaluator[r, ka, Unit] \ Read(r) = 
        let* cell = getCellAt(ix);
        if (Object.isNull(cell)) point() else throwError("getNullAt(${ix})")


    pub def getCellTypeAt(ix: Int32): RowEvaluator[r, ka, CellType] \ Read(r) =
        liftCellGetter(ix, (_, cell) -> {
            Interop/Apache/Poi/SS/Usermodel/Cell.getCellType(cell) 
                |> SheetIO/CellType.fromNative
        })


    pub def isBlankAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => point(true)
            case _       => point(false)
        } |> swapError("isBlankAt(${ix})")


    pub def getBlankAt(ix: Int32): RowEvaluator[r, ka, Unit] \ Read(r) = 
        let errMsg = "getBlankAt(${ix})";
        let* typ = getCellTypeAt(ix);
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)

    /// Error if the cell type is numeric.
    pub def getStringAt(ix: Int32): RowEvaluator[r, ka, String] \ Read(r) =
        liftCellGetterResult(ix, (_, row) -> Interop/Apache/Poi/SS/Usermodel/Cell.getStringCellValue(row))


    /// TODO - wrong name or should return Bool...
    pub def getSatifiesAt(ix: Int32, test: String -> Bool): RowEvaluator[r, ka, String] \ Read(r) =    
        let errMsg = "getSatifiesAt(${ix})";
        let step = () -> getContentAt(ix) >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)


    pub def getNonEmptyContentAt(ix: Int32): RowEvaluator[r, ka, String] \ Read(r) =
        getSatifiesAt(ix, x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getNonEmptyContentAt(${ix})")


    pub def getStringAtWithDefault(ix: Int32, d: String): RowEvaluator[r, ka, String] \ Read(r) =
        let* x = isNullAt(ix);
        let* y = isBlankAt(ix);
        if (x or y) point(d) else getStringAt(ix)


    pub def getStringAtWithDataFormatter(ix: Int32, dfmt: DataFormatter[r]): RowEvaluator[r, ka, String] \Read(r) =
        liftCellGetter(ix, (_, cell) -> Interop/Apache/Poi/SS/Usermodel/DataFormatter.formatCellValue(cell, dfmt))


    pub def getBoolAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) =
        liftCellGetterResult(ix, (_, cell) -> Interop/Apache/Poi/SS/Usermodel/Cell.getBooleanCellValue(cell))


    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumericAt(ix: Int32): RowEvaluator[r, ka, Float64] \ Read(r) =
        liftCellGetterResult(ix, (_, cell) -> Interop/Apache/Poi/SS/Usermodel/Cell.getNumericCellValue(cell))


    def getNumAt(ix: Int32, conv: Float64 -> Option[a]): RowEvaluator[r, ka, a] \ Read(r) =
        let* n = getNumericAt(ix);
        match conv(n) {
            case Some(i) => point(i)
            case None    => throwError("getNumAt")
        }


    pub def getInt8At(ix: Int32): RowEvaluator[r, ka, Int8] \ Read(r) =
        getNumAt(ix, Float64.tryToInt8) 
            |> swapError("getInt8At")

    pub def getInt16At(ix: Int32): RowEvaluator[r, ka, Int16] \ Read(r) =
        getNumAt(ix, Float64.tryToInt16) 
            |> swapError("getInt16At")

    pub def getInt32At(ix: Int32): RowEvaluator[r, ka, Int32] \ Read(r) =
        getNumAt(ix, Float64.tryToInt32) 
            |> swapError("getInt32At")

    pub def getInt64At(ix: Int32): RowEvaluator[r, ka, Int64] \ Read(r) =
        getNumAt(ix, Float64.tryToInt64) 
            |> swapError("getInt64At")

    pub def getFloat32At(ix: Int32): RowEvaluator[r, ka, Float32] \ Read(r) =
        getNumAt(ix, Float64.tryToFloat32) 
            |> swapError("getFloat32At")
    
    pub def getFloat64At(ix: Int32): RowEvaluator[r, ka, Float64] \ Read(r) = 
        getNumAt(ix, x -> Some(x)) 
            |> swapError("getFloat64At")


    def isDateFormattedAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        liftCellGetter(ix, ( _, cell) -> Interop/Apache/Poi/SS/Usermodel/DateUtil.isCellDateFormatted(cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    ///
    pub def getLocalDateTimeAt(ix: Int32): RowEvaluator[r, ka, Time.LocalDateTime] \ Read(r) = 
        use Time.LocalDateTime.{LocalDateTime};
        let* ctype = getCellTypeAt(ix);
        if (ctype == TyNumeric) {
            let* dtCell = isDateFormattedAt(ix);
            if (dtCell) {
                let* dt = liftCellGetterResult(ix, ( _, cell) -> Interop/Apache/Poi/SS/Usermodel/Cell.getLocalDateTimeCellValue(cell));
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeAt - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeAt - not a date")
        } else throwError("getLocalDateTimeAt - not a date / not a numeric cell")



    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithFormat(ix: Int32, 
                                        fmt: DateTimeFormatter[Time/Format.DateTimeFormat]): RowEvaluator[r, ka, Time.LocalDateTime] \ Read(r) = 
        let* s = getStringAt(ix);
        GradedError.liftResult(Time/LocalDateTime.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithPattern(ix: Int32, patt: String): RowEvaluator[r, ka, Time.LocalDateTime] \ Read(r) = 
        let* fmt    = Graded/MonadError.liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringAt(ix);
        GradedError.liftResult(Time/LocalDateTime.parse(fmt, s))

/*
    pub def getLocalDateAt(ix: Int32): RowEvaluator[r, ka, Time.LocalDate] \ Read(r) = 
        getLocalDateTimeAt(ix)
            |> map(Time/LocalDateTime.toLocalDate)
            |> swapError("getLocalDateAt(${ix})")



    /// Cell must be readable as a String
    pub def getLocalDateAtWithFormat(ix: Int32, 
                                    fmt: DateTimeFormatter[Time/Format.DateFormat]): RowEvaluator[r, ka, Time.LocalDate] \ Read(r) = 
        let* s = getStringAt(ix);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    /// Cell must be readable as a String
    pub def getLocalDateAtWithPattern(ix: Int32, patt: String): RowEvaluator[r, ka, Time.LocalDate] \ Read(r) = 
        let* fmt    = liftResult(Time/Format/DateTimeFormatter.fromPattern(patt));
        let* s      = getStringAt(ix);
        liftResult(Time/LocalDate.parse(fmt, s))
        

    def getIfNotEmpty(ix: Int32, getter: Int32 -> RowEvaluator[r, ka, a] \ Read(r)): RowEvaluator[r, ka, Option[a]] \ Read(r) = 
        let* x1 = isNullAt(ix);
        let* y1 = isBlankAt(ix);
        if (x1 or y1) point(None) else map(Some, getter(ix))

    pub def getOptionStringAt(ix: Int32): RowEvaluator[r, ka, Option[String]] \ Read(r) =
        getIfNotEmpty(ix, getStringAt)



    pub def getOptionInt8At(ix: Int32): RowEvaluator[r, ka, Option[Int8]] \ Read(r) =
        getIfNotEmpty(ix, getInt8At)

    pub def getOptionInt16At(ix: Int32): RowEvaluator[r, ka, Option[Int16]] \ Read(r) =
        getIfNotEmpty(ix, getInt16At)

    pub def getOptionInt32At(ix: Int32): RowEvaluator[r, ka, Option[Int32]] \ Read(r) =
        getIfNotEmpty(ix, getInt32At)

    pub def getOptionInt64At(ix: Int32): RowEvaluator[r, ka, Option[Int64]] \ Read(r) =
        getIfNotEmpty(ix, getInt64At)

    pub def getOptionFloat32At(ix: Int32): RowEvaluator[r, ka, Option[Float32]] \ Read(r) =
        getIfNotEmpty(ix, getFloat32At)
    
    pub def getOptionFloat64At(ix: Int32): RowEvaluator[r, ka, Option[Float64]] \ Read(r) = 
        getIfNotEmpty(ix, getFloat64At)

    pub def getOptionLocalDateTimeAt(ix: Int32): RowEvaluator[r, ka, Option[Time.LocalDateTime]] \ Read(r) = 
        getIfNotEmpty(ix, getLocalDateTimeAt)
*/
        
    pub def getFirstCellNum(): RowEvaluator[r, ka, Int32] \ Read(r) =
        liftGetter((_, row) -> Interop/Apache/Poi/SS/Usermodel/Row.getFirstCellNum(row) |> Int16.toInt32)

    pub def getLastCellNum(): RowEvaluator[r, ka, Int32] \ Read(r) =
        liftGetter((_, row) -> Interop/Apache/Poi/SS/Usermodel/Row.getLastCellNum(row) |> Int16.toInt32)
            |> map(n -> if (n > 0) n - 1 else n)

/*
    def queryFontAt(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] & ef): RowEvaluator[r, ka, a] \ { Read(r), ef } = 
        use Result.flatMap;
        liftPrim((r, _, row) -> 
            let* cell   = SheetIO/Internal/Row.getCell(r, ix, row);
            let* style  = SheetIO/Internal/Cell.getCellStyle(r, cell);
            let* ixf    = SheetIO/Internal/CellStyle.getFontIndexAsInt(r, style);
            let* sheet  = SheetIO/Internal/Row.getSheet(r, row);
            let* wb     = SheetIO/Internal/Sheet.getWorkbook(r, sheet);
            let* font   = SheetIO/Internal/Workbook.getFontAt(r, ixf, wb);
            let* ans    = fn(r, font); 
            Ok(ans)
        )




    pub def getIsBoldAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        queryFontAt(ix, SheetIO/Internal/Font.getBold)


    pub def getIsItalicAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        queryFontAt(ix, SheetIO/Internal/Font.getItalic)

    pub def getIsStrikeoutAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        queryFontAt(ix, SheetIO/Internal/Font.getStrikeout)

*/
    /// List (not Array) avoids region allocation...
    pub def getRowContents(): RowEvaluator[r, ka, List[String]] \ Read(r) =
        getRowContentsWith(getContentAt)


    pub def getRowContentsWith(evalCell: Int32 -> RowEvaluator[r, ka, a] \ ef): RowEvaluator[r, ka, List[a]] \ {Read(r), ef} =
        let* end    = getLastCellNum();
        getRowContentsWithHelper(evalCell, end, 0, ks -> unsafe_cast point(ks) as _ \ ef)


    // Should be tail recursive as bind is in the contiuation
    def getRowContentsWithHelper(evalCell: Int32 -> RowEvaluator[r, ka, a] \ ef, 
                                    end: Int32, 
                                    ix: Int32, 
                                    k: List[a] -> RowEvaluator[r, ka, List[a]] \ ef): RowEvaluator[r, ka, List[a]] \ ef =
        if (ix > end)
            k(Nil)
        else 
            getRowContentsWithHelper(evalCell, end, ix + 1, ss -> evalCell(ix) >>= s -> k(s :: ss))

/*
    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(r: Region[r2]): RowEvaluator[r, ka1, Array[String, r2]] \ { Read(r1), Write(r2) } =
        let eval = ix -> Graded/MonadOr.orelse(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(r, eval)

    def getIsoLocalDateTimeAsString(ix: Int32): RowEvaluator[r, ka, String] \ Read(r) = 
        let fmt = Time/Format/DateTimeFormatter.iso_local_date_time();
        let* dt = getLocalDateTimeAt(ix);
        let s   = Time/LocalDateTime.format(fmt, dt);
        point(s)

*/

}
