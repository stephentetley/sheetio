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

    use SheetIO.DataFormatter;
    
    /// TODO - add EnvEvaluator to GradedLib?

    // pass around DataFormatter[r]
    pub enum RowEvaluator[r: Region, ka: Type, ef: Bool, a: Type](
        Graded.EnvEvaluator[ka, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Row[r], ef, a])


    instance GradedFunctor[RowEvaluator[r, ka]] {
        pub def map(f: a -> b \ ef1, 
                    ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1, b] = 
            SheetIO.RowEvaluator.map(f, ma)
    }

    instance GradedApplicative[RowEvaluator[r, ka]] {
        pub def point(x: a): RowEvaluator[r, ka, ef, a] = 
            SheetIO.RowEvaluator.point(x)

        pub def ap(f: RowEvaluator[r, ka, ef1, a -> b \ ef2], 
                    ma: RowEvaluator[r, ka, ef, a]) : RowEvaluator[r, ka, ef and ef1 and ef2, b] = 
            SheetIO.RowEvaluator.ap(f, ma)
    }

    instance GradedMonad[RowEvaluator[r, ka]] {
        pub def flatMap(f: a -> RowEvaluator[r, ka, ef1, b] \ ef2, 
                        ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1 and ef2, b] =
            SheetIO.RowEvaluator.flatMap(f, ma)
    }

    instance GradedZero[RowEvaluator[r, ka]] {
        pub def empty(): RowEvaluator[r, ka, ef, a] = SheetIO.RowEvaluator.throwError("empty")
    }

    instance GradedOr[RowEvaluator[r, ka]] {
        pub def orelse(ma: RowEvaluator[r, ka, ef1, a], 
                        mb: RowEvaluator[r, ka, ef2, a]): RowEvaluator[r, ka, ef1 and ef2, a] =
            SheetIO.RowEvaluator.orelse(ma, mb)
    }

    instance GradedError[RowEvaluator[r, ka]] {
        pub def throwError(err: String): RowEvaluator[r, ka, ef, a] =
            SheetIO.RowEvaluator.throwError(err)

        pub def mapError(f: String -> String \ ef1, 
                            ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1, a] =
            SheetIO.RowEvaluator.mapError(f, ma)

        pub def catchError(ma: RowEvaluator[r, ka, ef, a], 
                            hdl: String -> RowEvaluator[r, ka, ef1, a] \ ef2): RowEvaluator[r, ka, ef and ef1 and ef2, a] =
            SheetIO.RowEvaluator.catchError(ma, hdl)

        pub def reifyError(ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef, Result[String, a]] =
            SheetIO.RowEvaluator.reifyError(ma)
    }

}

mod SheetIO.RowEvaluator {

    
    use GradedError.{swapError}
    use GradedMonad.{>>=}
    use GradedFunctor.{<$$>}
    use Time.Format.DateTimeFormatter
    
    use SheetIO.RowEvaluator
    use SheetIO.RowEvaluator.{RowEvaluator}

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter
    use Interop.Apache.Poi.SS.Usermodel.Row

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    

    pub def getRowEvaluator(x: RowEvaluator[r, ka, ef, a]): Graded.EnvEvaluator[ka, DataFormatter, Row[r], ef, a] = 
        let RowEvaluator(x1) = x;
        x1
    

    pub def runRowEvaluator(ma: RowEvaluator[r, a, ef, a], 
                            st: Row[r]): Result[String, a] \ ef =
        let env = ???;
        Graded.EnvEvaluator.runEnvEvaluator(getRowEvaluator(ma), env, st)
        

    pub def point(x: a): RowEvaluator[r, ka, ef, a] = 
        RowEvaluator(Graded.EnvEvaluator.point(x))
   
    pub def map(f: a -> b \ ef1, 
                ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1, b] =
        RowEvaluator(Graded.EnvEvaluator.map(f, getRowEvaluator(ma)))

    pub def ap(mf: RowEvaluator[r, ka, ef1, a -> b \ ef2], 
                ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1 and ef2, b] =
        RowEvaluator(Graded.EnvEvaluator.ap(getRowEvaluator(mf), getRowEvaluator(ma)))

    pub def flatMap(f: a -> RowEvaluator[r, ka, ef1, b] \ ef2, 
                    ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1 and ef2, b] =
        RowEvaluator(Graded.EnvEvaluator.flatMap(f >> getRowEvaluator, getRowEvaluator(ma)))

    pub def throwError(msg: String): RowEvaluator[r, ka, ef, a] = 
        RowEvaluator(Graded.EnvEvaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef1, ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and ef1, a] =
        RowEvaluator(Graded.EnvEvaluator.mapError(f, getRowEvaluator(ma)))

    pub def catchError(ma: RowEvaluator[r, ka, ef, a], 
                        f: String -> RowEvaluator[r, ka, ef1, a] \ ef2): RowEvaluator[r, ka, ef and ef1 and ef2, a] =
        RowEvaluator(Graded.EnvEvaluator.catchError(getRowEvaluator(ma), f >> getRowEvaluator))

    pub def reifyError(ma: RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef, Result[String, a]] =
        RowEvaluator(Graded.EnvEvaluator.reifyError(getRowEvaluator(ma)))

    pub def empty(): RowEvaluator[r, ka, ef, a] = 
        RowEvaluator(Graded.EnvEvaluator.empty())

    pub def orelse(ma: RowEvaluator[r, ka, ef1, a], mb: RowEvaluator[r, ka, ef2, a]): RowEvaluator[r, ka, ef1 and ef2, a] =
        RowEvaluator(Graded.EnvEvaluator.orelse(getRowEvaluator(ma), getRowEvaluator(mb)))
   
    /// `Graded.EnvEvaluator.liftGetter` not sufficiently polymorphic...
    def liftGetter(f: Row[r] -> a \ ef): RowEvaluator[r, ka, ef, a] =
        RowEvaluator(Graded.EnvEvaluator.liftGetter(f))
    
    /// `Graded.EnvEvaluator.liftGetterResult` not sufficiently polymorphic...
    def liftGetterResult(f: Row[r] -> Result[err, a] \ ef): RowEvaluator[r, ka, ef, a] with ToString[err] =
        RowEvaluator(Graded.EnvEvaluator.liftGetterResult(f))

    def _liftGetterWithEnv(f: DataFormatter -> Row[r] -> a \ ef): RowEvaluator[r, ka, ef, a] =
        RowEvaluator(Graded.EnvEvaluator.liftGetterWithEnv(f))
    
    def liftGetterResultWithEnv(f: DataFormatter -> Row[r] -> Result[err, a] \ ef): RowEvaluator[r, ka, ef, a] with ToString[err] =
        RowEvaluator(Graded.EnvEvaluator.liftGetterResultWithEnv(f))



    pub def liftAction(f: Unit -> a \ ef): RowEvaluator[r, ka, ef, a] =
        RowEvaluator(Graded.EnvEvaluator.liftAction(f))
    
        
    // pub def liftActionResult(f: Unit -> Result[err, a] \ ef): RowEvaluator[r, ka, a] \ ef with ToString[err] =
    //     RowEvaluator(Graded.EnvEvaluator.liftActionResult(f))



    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    def getCellAt(ix: Int32): RowEvaluator[r, ka, r, Cell[r]] = 
        liftGetterResult(row -> Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) |> Option.toOk("getCellAt(${ix})"))


    def liftCellGetter(ix: Int32, 
                        fn: DataFormatter -> Cell[r] -> a \ ef): RowEvaluator[r, ka, ef and r, a] =
        liftGetterResultWithEnv((fmt, row) -> 
            match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                case Some(cell) => fn(fmt, cell) |> Ok
                case None       => Err("liftCellGetter(${ix})")
            }
        )

    def liftCellGetterResult(ix: Int32, 
                            fn: DataFormatter -> Cell[r] -> Result[err, a] \ ef): RowEvaluator[r, ka, ef and r, a] with ToString[err] =
        liftGetterResultWithEnv((fmt, row) -> 
            match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                case Some(cell) => fn(fmt, cell) |> Result.mapErr(ToString.toString)
                case None       => Err("liftCellGetterResult(${ix})")
            }
        )

    pub def getRowNum(): RowEvaluator[r, ka, r, Int32] = 
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row))


    pub def getSheetName(): RowEvaluator[r, ka, r, String] = 
        liftGetter(row -> {
            let sheet = Interop.Apache.Poi.SS.Usermodel.Row.getSheet(row);
            Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName(sheet)
        })
            

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    /// Effect not sufficiently polymorphic...
    pub def getColumnIndex(s: String): RowEvaluator[r, ka, ef, Int32] = 
        liftAction(_ -> 
            let addr = Interop.Apache.Poi.SS.Util.CellAddress.cellAddressWithAddress(s + "1");
            checked_ecast(Interop.Apache.Poi.SS.Util.CellAddress.getColumn(addr))
        )


    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowEvaluator[r, ka, r, String] =
        liftCellGetter(ix, (fmt, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, fmt))


    /// Uses the default DataFormatter
    pub def getTrimmedContentAt(ix: Int32): RowEvaluator[r, ka, r, String] =
        map(String.trim, getContentAt(ix))

    pub def isNullAt(ix: Int32): RowEvaluator[r, ka, r, Bool] = 
        getCellAt(ix) <$$> Object.isNull

    pub def getNullAt(ix: Int32): RowEvaluator[r, ka, r, Unit] = 
        getCellAt(ix)   >>= cell -> 
        if (Object.isNull(cell)) point() else throwError("getNullAt(${ix})")


    pub def getCellTypeAt(ix: Int32): RowEvaluator[r, ka, r, CellType] =
        liftCellGetter(ix, (_, cell) -> {
            Interop.Apache.Poi.SS.Usermodel.Cell.getCellType(cell) 
                |> SheetIO.CellType.fromNative
        })


    pub def isBlankAt(ix: Int32): RowEvaluator[r, ka, r, Bool] = 
        getCellTypeAt(ix) >>= typ ->
        match typ { 
            case TyBlank => point(true)
            case _       => point(false)
        } |> swapError("isBlankAt(${ix})")


    pub def getBlankAt(ix: Int32): RowEvaluator[r, ka, r, Unit] = 
        let errMsg = "getBlankAt(${ix})";
        getCellTypeAt(ix) >>= typ ->
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)

    /// Error if the cell type is numeric.
    pub def getStringAt(ix: Int32): RowEvaluator[r, ka, r, String] =
        liftCellGetterResult(ix, (_, row) -> Interop.Apache.Poi.SS.Usermodel.Cell.getStringCellValue(row))


    /// TODO - wrong name or should return Bool...
    pub def getSatifiesAt(ix: Int32, test: String -> Bool): RowEvaluator[r, ka, r, String] =
        let errMsg = "getSatifiesAt(${ix})";
        let step = () -> getContentAt(ix) >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)


    pub def getNonEmptyContentAt(ix: Int32): RowEvaluator[r, ka, r, String] =
        getSatifiesAt(ix, x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getNonEmptyContentAt(${ix})")


    pub def getStringAtWithDefault(ix: Int32, d: String): RowEvaluator[r, ka, r, String]=
        isNullAt(ix)    >>= nullc -> 
        isBlankAt(ix)   >>= blankc -> 
        if (nullc or blankc) point(d) else getStringAt(ix)


    pub def getStringAtWithDataFormatter(ix: Int32, dfmt: DataFormatter): RowEvaluator[r, ka, r, String] =
        liftCellGetter(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, dfmt))


    pub def getBoolAt(ix: Int32): RowEvaluator[r, ka, r, Bool] =
        liftCellGetterResult(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getBooleanCellValue(cell))

    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumericAt(ix: Int32): RowEvaluator[r, ka, r, Float64] =
        liftCellGetterResult(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.getNumericCellValue(cell))


    def getNumAt(ix: Int32, conv: Float64 -> Option[a]): RowEvaluator[r, ka, r, a] =
        getNumericAt(ix) >>= n -> 
        match conv(n) {
            case Some(i) => point(i)
            case None    => throwError("getNumAt")
        }


    pub def getInt8At(ix: Int32): RowEvaluator[r, ka, r, Int8] =
        getNumAt(ix, Float64.tryToInt8) 
            |> swapError("getInt8At")

    pub def getInt16At(ix: Int32): RowEvaluator[r, ka, r, Int16] =
        getNumAt(ix, Float64.tryToInt16) 
            |> swapError("getInt16At")

    pub def getInt32At(ix: Int32): RowEvaluator[r, ka, r, Int32] =
        getNumAt(ix, Float64.tryToInt32) 
            |> swapError("getInt32At")

    pub def getInt64At(ix: Int32): RowEvaluator[r, ka, r, Int64] =
        getNumAt(ix, Float64.tryToInt64) 
            |> swapError("getInt64At")

    pub def getFloat32At(ix: Int32): RowEvaluator[r, ka, r, Float32] =
        getNumAt(ix, Float64.tryToFloat32) 
            |> swapError("getFloat32At")
    
    pub def getFloat64At(ix: Int32): RowEvaluator[r, ka, r, Float64] = 
        getNumAt(ix, x -> Some(x)) 
            |> swapError("getFloat64At")


    def isDateFormattedAt(ix: Int32): RowEvaluator[r, ka, r, Bool] = 
        liftCellGetter(ix, ( _, cell) -> Interop.Apache.Poi.SS.Usermodel.DateUtil.isCellDateFormatted(cell))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com.questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    ///
    pub def getLocalDateTimeAt(ix: Int32): RowEvaluator[r, ka, r, Time.LocalDateTime] = 
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
                                        fmt: DateTimeFormatter[Time.Format.DateTimeFormat]): RowEvaluator[r, ka, r, Time.LocalDateTime] = 
        getStringAt(ix) >>= s ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))


    /// Cell must be readable as a String
    pub def getLocalDateTimeAtWithPattern(ix: Int32, patt: String): RowEvaluator[r, ka, r, Time.LocalDateTime] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getStringAt(ix)                                                         >>= s   ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))


    pub def getLocalDateAt(ix: Int32): RowEvaluator[r, ka, r, Time.LocalDate] = 
        getLocalDateTimeAt(ix)
            |> map(Time.LocalDateTime.toLocalDate)
            |> swapError("getLocalDateAt(${ix})")


    /// Cell must be readable as a String
    pub def getLocalDateAtWithFormat(ix: Int32, 
                                    fmt: DateTimeFormatter[Time.Format.DateFormat]): RowEvaluator[r, ka, r, Time.LocalDate] = 
        getStringAt(ix) >>= s -> 
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))


    /// Cell must be readable as a String
    pub def getLocalDateAtWithPattern(ix: Int32, patt: String): RowEvaluator[r, ka, r, Time.LocalDate] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getStringAt(ix)                                                         >>= s   ->
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))
        

    def getIfNotEmpty(ix: Int32, getter: Int32 -> RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and r, Option[a]] = 
        isNullAt(ix)    >>= nullc   ->
        isBlankAt(ix)   >>= blankc  ->
        if (nullc or blankc) point(None) else map(Some, getter(ix))

    pub def getOptionStringAt(ix: Int32): RowEvaluator[r, ka, r, Option[String]] =
        getIfNotEmpty(ix, getStringAt)



    pub def getOptionInt8At(ix: Int32): RowEvaluator[r, ka, r, Option[Int8]] =
        getIfNotEmpty(ix, getInt8At)

    pub def getOptionInt16At(ix: Int32): RowEvaluator[r, ka, r, Option[Int16]] =
        getIfNotEmpty(ix, getInt16At)

    pub def getOptionInt32At(ix: Int32): RowEvaluator[r, ka, r, Option[Int32]] =
        getIfNotEmpty(ix, getInt32At)

    pub def getOptionInt64At(ix: Int32): RowEvaluator[r, ka, r, Option[Int64]] =
        getIfNotEmpty(ix, getInt64At)

    pub def getOptionFloat32At(ix: Int32): RowEvaluator[r, ka, r, Option[Float32]] =
        getIfNotEmpty(ix, getFloat32At)
    
    pub def getOptionFloat64At(ix: Int32): RowEvaluator[r, ka, r, Option[Float64]] = 
        getIfNotEmpty(ix, getFloat64At)

    pub def getOptionLocalDateTimeAt(ix: Int32): RowEvaluator[r, ka, r, Option[Time.LocalDateTime]] = 
        getIfNotEmpty(ix, getLocalDateTimeAt)
      
    pub def getFirstCellNum(): RowEvaluator[r, ka, r, Int32] =
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getFirstCellNum(row) |> Int16.toInt32)

    pub def getLastCellNum(): RowEvaluator[r, ka, r, Int32] =
        liftGetter(row -> Interop.Apache.Poi.SS.Usermodel.Row.getLastCellNum(row) |> Int16.toInt32)
            |> map(n -> if (n > 0) n - 1 else n)


/*
    def queryFontAt(ix: Int32, 
                    fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] \ ef): RowEvaluator[r, ka, ef and r, a] = 
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

    pub def getIsBoldAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        queryFontAt(ix, SheetIO.Internal.Font.getBold)


    pub def getIsItalicAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        queryFontAt(ix, SheetIO.Internal.Font.getItalic)

    pub def getIsStrikeoutAt(ix: Int32): RowEvaluator[r, ka, Bool] \ Read(r) = 
        queryFontAt(ix, SheetIO.Internal.Font.getStrikeout)
*/

    /// List (not Array) avoids region allocation...
    pub def getRowContents(): RowEvaluator[r, ka, r, Vector[String]] =
        getRowContentsWith(getContentAt)


    pub def getRowContentsWith(evalCell: Int32 -> RowEvaluator[r, ka, ef, a]): RowEvaluator[r, ka, ef and r, Vector[a]] =
        getLastCellNum()                                            >>= end ->
        getRowContentsWithHelper(evalCell, end, 0, ks -> point(ks)) >>= xs  ->
        point(List.toVector(xs))


    // Should be tail recursive as bind is in the contiuation
    def getRowContentsWithHelper(evalCell: Int32 -> RowEvaluator[r, ka, ef, a], 
                                    end: Int32, 
                                    ix: Int32, 
                                    k: List[a] -> RowEvaluator[r, ka, ef, List[a]]): RowEvaluator[r, ka, ef, List[a]] =
        if (ix > end)
            k(Nil)
        else 
            getRowContentsWithHelper(evalCell, end, ix + 1, ss -> evalCell(ix) >>= s -> k(s :: ss))

    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(): RowEvaluator[r, ka, r, Vector[String]]  =
        let eval = ix -> GradedOr.orelse(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(eval)

    def getIsoLocalDateTimeAsString(ix: Int32): RowEvaluator[r, ka, r, String] = 
        let fmt = Time.Format.DateTimeFormatter.iso_local_date_time();
        getLocalDateTimeAt(ix)                              >>= dt  ->
        point(Time.LocalDateTime.format(fmt, dt))



}
