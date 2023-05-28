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

    pub type alias RowEvaluator[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.EnvEvaluator[ka, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Row[r], ef, a]

}

mod SheetIO.RowEvaluator {

    
    use GradedFunctor.{map, <$$>}
    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{throwError, swapError}
    use Graded.EnvEvaluator.{liftGetter, liftGetterResult, liftGetterResultWithEnv, liftAction}
    use Time.Format.DateTimeFormatter
    
    use SheetIO.RowEvaluator

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter
    use Interop.Apache.Poi.SS.Usermodel.Row

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    


    pub def runRowEvaluator(ma: RowEvaluator[a, r, ef, a], 
                            env: DataFormatter,
                            st: Row[r]): Result[String, a] \ ef =
        Graded.EnvEvaluator.runEnvEvaluator(ma, env, st)
        



    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    def getCellAt(ix: Int32): RowEvaluator[ka, r, r, Cell[r]] = 
        liftGetterResult(row -> Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) |> Option.toOk("getCellAt(${ix})"))


    def liftCellGetter(ix: Int32, 
                        fn: DataFormatter -> Cell[r] -> a \ ef): RowEvaluator[ka, r, ef + r, a] =
        liftGetterResultWithEnv((fmt, row) -> 
            match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                case Some(cell) => fn(fmt, cell) |> Ok
                case None       => Err("liftCellGetter(${ix})")
            }
        )

    def liftCellGetterResult(ix: Int32, 
                            fn: DataFormatter -> Cell[r] -> Result[err, a] \ ef): RowEvaluator[ka, r, ef + r, a] with ToString[err] =
        liftGetterResultWithEnv((fmt, row) -> 
            match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                case Some(cell) => fn(fmt, cell) |> Result.mapErr(ToString.toString)
                case None       => Err("liftCellGetterResult(${ix})")
            }
        )

    pub def getRowNum(): RowEvaluator[ka, r, r, Int32] = 
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row))


    pub def getSheetName(): RowEvaluator[ka, r, r, String] = 
        liftGetter(row -> {
            let sheet = Interop.Apache.Poi.SS.Usermodel.Row.getSheet(row);
            Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName(sheet)
        })
            

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    /// Effect not sufficiently polymorphic...
    pub def getColumnIndex(s: String): RowEvaluator[ka, r, ef, Int32] = 
        liftAction(_ -> 
            let addr = Interop.Apache.Poi.SS.Util.CellAddress.cellAddressWithAddress(s + "1");
            checked_ecast(Interop.Apache.Poi.SS.Util.CellAddress.getColumn(addr))
        )


    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowEvaluator[ka, r, r, String] =
        liftCellGetter(ix, (fmt, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, fmt))


    /// Uses the default DataFormatter
    pub def getTrimmedContentAt(ix: Int32): RowEvaluator[ka, r, r, String] =
        map(String.trim, getContentAt(ix))

    pub def isNullAt(ix: Int32): RowEvaluator[ka, r, r, Bool] = 
        getCellAt(ix) <$$> Object.isNull

    pub def getNullAt(ix: Int32): RowEvaluator[ka, r, r, Unit] = 
        getCellAt(ix)   >>= cell -> 
        if (Object.isNull(cell)) point() else throwError("getNullAt(${ix})")


    pub def getCellTypeAt(ix: Int32): RowEvaluator[ka, r, r, CellType] =
        liftCellGetter(ix, (_, cell) -> {
            Interop.Apache.Poi.SS.Usermodel.Cell.getCellType(cell) 
                |> SheetIO.CellType.fromNative
        })


    pub def isBlankAt(ix: Int32): RowEvaluator[ka, r, r, Bool] = 
        getCellTypeAt(ix) >>= typ ->
        match typ { 
            case TyBlank => point(true)
            case _       => point(false)
        } |> swapError("isBlankAt(${ix})")


    pub def getBlankAt(ix: Int32): RowEvaluator[ka, r, r, Unit] = 
        let errMsg = "getBlankAt(${ix})";
        getCellTypeAt(ix) >>= typ ->
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)

    /// Error if the cell type is numeric.
    pub def getStringAt(ix: Int32): RowEvaluator[ka, r, r, String] =
        liftCellGetterResult(ix, (_, row) -> Interop.Apache.Poi.SS.Usermodel.Cell.getStringCellValue(row))


    /// TODO - wrong name or should return Bool...
    pub def getSatifiesAt(ix: Int32, test: String -> Bool): RowEvaluator[ka, r, r, String] =
        let errMsg = "getSatifiesAt(${ix})";
        let step = () -> getContentAt(ix) >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)


    pub def getNonEmptyContentAt(ix: Int32): RowEvaluator[ka, r, r, String] =
        getSatifiesAt(ix, x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getNonEmptyContentAt(${ix})")


    pub def getStringAtWithDefault(ix: Int32, d: String): RowEvaluator[ka, r, r, String]=
        isNullAt(ix)    >>= nullc -> 
        isBlankAt(ix)   >>= blankc -> 
        if (nullc or blankc) point(d) else getStringAt(ix)


    pub def getStringAtWithDataFormatter(ix: Int32, dfmt: DataFormatter): RowEvaluator[ka, r, r, String] =
        liftCellGetter(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, dfmt))


    pub def getBoolAt(ix: Int32): RowEvaluator[ka, r, r, Bool] =
        liftCellGetterResult(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getBooleanCellValue(cell))

    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumericAt(ix: Int32): RowEvaluator[ka, r, r, Float64] =
        liftCellGetterResult(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getNumericCellValue(cell))


    def getNumAt(ix: Int32, conv: Float64 -> Option[a]): RowEvaluator[ka, r, r, a] =
        getNumericAt(ix) >>= n -> 
        match conv(n) {
            case Some(i) => point(i)
            case None    => throwError("getNumAt")
        }


    pub def getInt8At(ix: Int32): RowEvaluator[ka, r, r, Int8] =
        getNumAt(ix, Float64.tryToInt8) 
            |> swapError("getInt8At")

    pub def getInt16At(ix: Int32): RowEvaluator[ka, r, r, Int16] =
        getNumAt(ix, Float64.tryToInt16) 
            |> swapError("getInt16At")

    pub def getInt32At(ix: Int32): RowEvaluator[ka, r, r, Int32] =
        getNumAt(ix, Float64.tryToInt32) 
            |> swapError("getInt32At")

    pub def getInt64At(ix: Int32): RowEvaluator[ka, r, r, Int64] =
        getNumAt(ix, Float64.tryToInt64) 
            |> swapError("getInt64At")

    pub def getFloat32At(ix: Int32): RowEvaluator[ka, r, r, Float32] =
        getNumAt(ix, Float64.tryToFloat32) 
            |> swapError("getFloat32At")
    
    pub def getFloat64At(ix: Int32): RowEvaluator[ka, r, r, Float64] = 
        getNumAt(ix, x -> Some(x)) 
            |> swapError("getFloat64At")


    def isDateFormattedAt(ix: Int32): RowEvaluator[ka, r, r, Bool] = 
        liftCellGetter(ix, ( _, cell) -> Interop.Apache.Poi.SS.Usermodel.DateUtil.isCellDateFormatted(cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com.questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    ///
    pub def getLocalDateTimeAt(ix: Int32): RowEvaluator[ka, r, r, Time.LocalDateTime] = 
        use Time.LocalDateTime.{LocalDateTime};
        getCellTypeAt(ix) >>= ctype ->
        if (ctype == TyNumeric) {
            isDateFormattedAt(ix) >>= dtCell ->
            if (dtCell) {
                liftCellGetterResult(ix, ( _, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getLocalDateTimeCellValue(cell)) >>= dt ->
                if (Object.isNull(dt))
                    throwError("getLocalDateTimeAt - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTimeAt - not a date")
        } else throwError("getLocalDateTimeAt - not a date / not a numeric cell")



    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithFormat(ix: Int32, 
                                        fmt: DateTimeFormatter[Time.Format.DateTimeFormat]): RowEvaluator[ka, r, r, Time.LocalDateTime] = 
        getStringAt(ix) >>= s ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))


    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithPattern(ix: Int32, patt: String): RowEvaluator[ka, r, r, Time.LocalDateTime] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getStringAt(ix)                                                         >>= s   ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))


    pub def getLocalDateAt(ix: Int32): RowEvaluator[ka, r, r, Time.LocalDate] = 
        getLocalDateTimeAt(ix)
            |> map(Time.LocalDateTime.toLocalDate)
            |> swapError("getLocalDateAt(${ix})")


    /// Cell must be readable as a String
    pub def getLocalDateAtWithFormat(ix: Int32, 
                                    fmt: DateTimeFormatter[Time.Format.DateFormat]): RowEvaluator[ka, r, r, Time.LocalDate] = 
        getStringAt(ix) >>= s -> 
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))


    /// Cell must be readable as a String
    pub def getLocalDateAtWithPattern(ix: Int32, patt: String): RowEvaluator[ka, r, r, Time.LocalDate] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getStringAt(ix)                                                         >>= s   ->
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))
        

    def getIfNotEmpty(ix: Int32, getter: Int32 -> RowEvaluator[ka, r, ef, a]): RowEvaluator[ka, r, ef + r, Option[a]] = 
        isNullAt(ix)    >>= nullc   ->
        isBlankAt(ix)   >>= blankc  ->
        if (nullc or blankc) point(None) else map(Some, getter(ix))

    pub def getOptionStringAt(ix: Int32): RowEvaluator[ka, r, r, Option[String]] =
        getIfNotEmpty(ix, getStringAt)



    pub def getOptionInt8At(ix: Int32): RowEvaluator[ka, r, r, Option[Int8]] =
        getIfNotEmpty(ix, getInt8At)

    pub def getOptionInt16At(ix: Int32): RowEvaluator[ka, r, r, Option[Int16]] =
        getIfNotEmpty(ix, getInt16At)

    pub def getOptionInt32At(ix: Int32): RowEvaluator[ka, r, r, Option[Int32]] =
        getIfNotEmpty(ix, getInt32At)

    pub def getOptionInt64At(ix: Int32): RowEvaluator[ka, r, r, Option[Int64]] =
        getIfNotEmpty(ix, getInt64At)

    pub def getOptionFloat32At(ix: Int32): RowEvaluator[ka, r, r, Option[Float32]] =
        getIfNotEmpty(ix, getFloat32At)
    
    pub def getOptionFloat64At(ix: Int32): RowEvaluator[ka, r, r, Option[Float64]] = 
        getIfNotEmpty(ix, getFloat64At)

    pub def getOptionLocalDateTimeAt(ix: Int32): RowEvaluator[ka, r, r, Option[Time.LocalDateTime]] = 
        getIfNotEmpty(ix, getLocalDateTimeAt)
      
    pub def getFirstCellNum(): RowEvaluator[ka, r, r, Int32] =
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getFirstCellNum(row) |> Int16.toInt32)

    pub def getLastCellNum(): RowEvaluator[ka, r, r, Int32] =
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getLastCellNum(row) |> Int16.toInt32)
            |> map(n -> if (n > 0) n - 1 else n)


/*
    def queryFontAt(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] \ ef): RowEvaluator[ka, r, ef + r, a] = 
        use Result.flatMap;
        liftPrim((r, _, row) -> 
            let* cell   = SheetIO.Internal.Row.getCell(r, ix, row);
            let* style  = SheetIO.Internal.Cell.getCellStyle(r, cell);
            let* ixf    = SheetIO.Internal.CellStyle.getFontIndexAsInt(r, style);
            let* sheet  = SheetIO.Internal.Row.getSheet(r, row);
            let* wb     = SheetIO.Internal.Sheet.getWorkbook(r, sheet);
            let* font   = SheetIO.Internal.Workbook.getFontAt(r, ixf, wb);
            let* ans    = fn(r, font); 
            Ok(ans)
        )

    pub def getIsBoldAt(ix: Int32): RowEvaluator[ka, r, Bool] \ r = 
        queryFontAt(ix, SheetIO.Internal.Font.getBold)


    pub def getIsItalicAt(ix: Int32): RowEvaluator[ka, r, Bool] \ r = 
        queryFontAt(ix, SheetIO.Internal.Font.getItalic)

    pub def getIsStrikeoutAt(ix: Int32): RowEvaluator[ka, r, Bool] \ r = 
        queryFontAt(ix, SheetIO.Internal.Font.getStrikeout)
*/

    /// List (not Array) avoids region allocation...
    pub def getRowContents(): RowEvaluator[ka, r, r, Vector[String]] =
        getRowContentsWith(getContentAt)


    pub def getRowContentsWith(evalCell: Int32 -> RowEvaluator[ka, r, ef, a]): RowEvaluator[ka, r, ef + r, Vector[a]] =
        getLastCellNum()                                            >>= end ->
        getRowContentsWithHelper(evalCell, end, 0, ks -> point(ks)) >>= xs  ->
        point(List.toVector(xs))


    // Should be tail recursive as bind is in the contiuation
    def getRowContentsWithHelper(evalCell: Int32 -> RowEvaluator[ka, r, ef, a], 
                                    end: Int32, 
                                    ix: Int32, 
                                    k: List[a] -> RowEvaluator[ka, r, ef, List[a]]): RowEvaluator[ka, r, ef, List[a]] =
        if (ix > end)
            k(Nil)
        else 
            getRowContentsWithHelper(evalCell, end, ix + 1, ss -> evalCell(ix) >>= s -> k(s :: ss))

    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(): RowEvaluator[ka, r, r, Vector[String]]  =
        let eval = ix -> GradedOr.orelse(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(eval)

    def getIsoLocalDateTimeAsString(ix: Int32): RowEvaluator[ka, r, r, String] = 
        let fmt = Time.Format.DateTimeFormatter.iso_local_date_time();
        getLocalDateTimeAt(ix)                              >>= dt  ->
        point(Time.LocalDateTime.format(fmt, dt))



}
