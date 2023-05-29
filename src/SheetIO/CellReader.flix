/*
 * Copyright 2023 Stephen Tetley
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

    pub type alias CellReader[ka: Type, ef: Eff, a: Type] = 
        Graded.EnvEvaluator[ka, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Cell, ef, a]


}

mod SheetIO.CellReader {

    use GradedFunctor.{map, <$$>}
    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{throwError, swapError}
    use Graded.EnvEvaluator.{liftGetter, liftGetterResult, liftGetterWithEnv}

    use Time.Format.DateTimeFormatter

    use SheetIO.CellReader
    use SheetIO.Cell
    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter

    pub def runCellEvaluator(ma: CellReader[a, ef, a], 
                            env: DataFormatter,
                            st: Cell): Result[String, a] \ ef =
        Graded.EnvEvaluator.runEnvEvaluator(ma, env, st)


    /// Uses the default DataFormatter
    pub def getContent(): CellReader[ka, ef, String] =
        liftGetterWithEnv((fmt, cell) -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, fmt)))

    /// Uses the default DataFormatter
    pub def getTrimmedContent(): CellReader[ka, ef, String] =
        getContent() <$$> String.trim

    pub def isNull(): CellReader[ka, ef, Bool] = 
        liftGetter(cell -> checked_ecast(Object.isNull(cell)))

    pub def getCellType(): CellReader[ka, ef, CellType] =
        liftGetter(cell -> 
            checked_ecast(Interop.Apache.Poi.SS.Usermodel.Cell.getCellType(cell))
                |> SheetIO.CellType.fromNative
        )

    pub def isBlank(): CellReader[ka, ef, Bool] = 
        (getCellType() <$$> typ -> typ == TyBlank) |> swapError("isBlank")


    pub def getBlank(): CellReader[ka, ef, Unit] = 
        (getCellType() >>= typ ->
            match typ { 
                case TyBlank => point()
                case _       => throwError("getBlank")
            }) |> swapError("getBlank")

    /// Error if the cell type is numeric.
    pub def getString(): CellReader[ka, ef, String] =
        liftGetterResult(cell -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Cell.getStringCellValue(cell)))


    /// TODO - wrong name or should return Bool...
    pub def getSatifies(test: String -> Bool): CellReader[ka, ef, String] =
        let errMsg = "getSatifies";
        let step = () -> getContent() >>= x -> 
                    if (test(x)) point(x) else throwError(errMsg);
        step() |> swapError(errMsg)

    pub def getNonEmptyContent(): CellReader[ka, ef, String] =
        getSatifies(x -> not String.isEmpty(String.trim(x)))
            |> map(String.trim)
            |> swapError("getNonEmptyContent")


    pub def getStringWithDefault(d: String): CellReader[ka, ef, String]=
        isNull()      >>= nullc -> 
        isBlank()     >>= blankc -> 
        if (nullc or blankc) point(d) else getString()


    pub def getStringWithDataFormatter(df: DataFormatter): CellReader[ka, ef, String] =
        liftGetter(cell -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, df)))


    pub def getBool(): CellReader[ka, ef, Bool] =
        liftGetterResult(cell -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Cell.getBooleanCellValue(cell)))

    /// Error if the cell type is String or cell value is not a readable double.
    /// Returns 0 for blank cells.
    pub def getNumeric(): CellReader[ka, ef, Float64] =
        liftGetterResult(cell -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Cell.getNumericCellValue(cell)))


    def getNum(conv: Float64 -> Option[a]): CellReader[ka, ef, a] =
        getNumeric()        >>= n -> 
        match conv(n) {
            case Some(i) => point(i)
            case None    => throwError("getNum")
        }


    pub def getInt8(): CellReader[ka, ef, Int8] =
        getNum(Float64.tryToInt8) 
            |> swapError("getInt8")

    pub def getInt16(): CellReader[ka, ef, Int16] =
        getNum(Float64.tryToInt16) 
            |> swapError("getInt16")

    pub def getInt32(): CellReader[ka, ef, Int32] =
        getNum(Float64.tryToInt32) 
            |> swapError("getInt32")

    pub def getInt64(): CellReader[ka, ef, Int64] =
        getNum(Float64.tryToInt64) 
            |> swapError("getInt64")

    pub def getFloat32(): CellReader[ka, ef, Float32] =
        getNum(Float64.tryToFloat32) 
            |> swapError("getFloat32")
    
    pub def getFloat64(): CellReader[ka, ef, Float64] = 
        getNum(x -> Some(x)) 
            |> swapError("getFloat64")


    def isDateFormatted(): CellReader[ka, ef, Bool] = 
        liftGetter(cell -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.DateUtil.isCellDateFormatted(cell)))

    /// This tests for cellType = NUMERIC and `isCellDateFormatted` before reading
    /// See https://stackoverflow.com.questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
    ///
    pub def getLocalDateTime(): CellReader[ka, ef, Time.LocalDateTime] = 
        use Time.LocalDateTime.{LocalDateTime};
        getCellType() >>= ctype ->
        if (ctype == TyNumeric) {
            isDateFormatted()       >>= dtCell ->
            if (dtCell) {
                liftGetterResult(cell -> Interop.Apache.Poi.SS.Usermodel.Cell.getLocalDateTimeCellValue(cell)) >>= dt ->
                if (Object.isNull(dt))
                    throwError("getLocalDateTime - null")
                else
                    point(LocalDateTime(dt))
            } else throwError("getLocalDateTime - not a date")
        } else throwError("getLocalDateTime - not a date / not a numeric cell")



    /// Cell must be readable as a String
    pub def getLocalDateTimeWithFormat(fmt: DateTimeFormatter[Time.Format.DateTimeFormat]): CellReader[ka, ef, Time.LocalDateTime] = 
        getString() >>= s ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))


    /// Cell must be readable as a String
    pub def getLocalDateTimeWithPattern(patt: String): CellReader[ka, ef, Time.LocalDateTime] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getString()                                                             >>= s   ->
        GradedError.liftResult(Time.LocalDateTime.parse(fmt, s))


    pub def getLocalDate(): CellReader[ka, ef, Time.LocalDate] = 
        getLocalDateTime()
            |> map(Time.LocalDateTime.toLocalDate)
            |> swapError("getLocalDate")


    /// Cell must be readable as a String
    pub def getLocalDateWithFormat(fmt: DateTimeFormatter[Time.Format.DateFormat]): CellReader[ka, ef, Time.LocalDate] = 
        getString() >>= s -> 
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))


    /// Cell must be readable as a String
    pub def getLocalDateWithPattern(patt: String): CellReader[ka, ef, Time.LocalDate] = 
        GradedError.liftResult(Time.Format.DateTimeFormatter.fromPattern(patt)) >>= fmt ->
        getString()                                                             >>= s   ->
        GradedError.liftResult(Time.LocalDate.parse(fmt, s))

    def getIfNotEmpty(getter: CellReader[ka, ef, a]): CellReader[ka, ef, Option[a]] = 
        isNull()        >>= nullc   ->
        isBlank()       >>= blankc  ->
        if (nullc or blankc) point(None) else map(Some, getter)


    pub def getOptionString(): CellReader[ka, ef, Option[String]] =
        getIfNotEmpty(getString())

    pub def getOptionInt8(): CellReader[ka, ef, Option[Int8]] =
        getIfNotEmpty(getInt8())

    pub def getOptionInt16(): CellReader[ka, ef, Option[Int16]] =
        getIfNotEmpty(getInt16())

    pub def getOptionInt32(): CellReader[ka, ef, Option[Int32]] =
        getIfNotEmpty(getInt32())

    pub def getOptionInt64(): CellReader[ka, ef, Option[Int64]] =
        getIfNotEmpty(getInt64())

    pub def getOptionFloat32(): CellReader[ka, ef, Option[Float32]] =
        getIfNotEmpty(getFloat32())
    
    pub def getOptionFloat64(): CellReader[ka, ef, Option[Float64]] = 
        getIfNotEmpty(getFloat64())

    pub def getOptionLocalDateTime(): CellReader[ka, ef, Option[Time.LocalDateTime]] = 
        getIfNotEmpty(getLocalDateTime())


/*
    def queryFont(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Font -> Result[a, PrimError] \ ef): CellReader[ka, r, ef + r, a] = 
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

    pub def getIsBold(): CellReader[ka, r, Bool] \ r = 
        queryFontAt(SheetIO.Internal.Font.getBold)


    pub def getIsItalic(): CellReader[ka, r, Bool] \ r = 
        queryFontAt(SheetIO.Internal.Font.getItalic)

    pub def getIsStrikeout(): CellReader[ka, r, Bool] \ r = 
        queryFontAt(SheetIO.Internal.Font.getStrikeout)

*/

}