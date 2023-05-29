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

    pub type alias RowReader[ka: Type, ef: Eff, a: Type] = 
        Graded.EnvEvaluator[ka, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Row, ef, a]

}

mod SheetIO.RowReader {

    
    use GradedFunctor.{map, <$$>}
    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{throwError, swapError}
    use Graded.EnvEvaluator.{liftGetter, liftGetterResult, liftGetterResultWithEnv, liftAction}
    use Time.Format.DateTimeFormatter
    
    use SheetIO.RowReader

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter
    use Interop.Apache.Poi.SS.Usermodel.Row

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    


    pub def runRowEvaluator(ma: RowReader[a, ef, a], 
                            env: DataFormatter,
                            st: Row): Result[String, a] \ ef =
        Graded.EnvEvaluator.runEnvEvaluator(ma, env, st)
        



    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.

    pub def getCellAt(ix: Int32): RowReader[ka, ef, Cell] = 
        liftGetterResult(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row)) |> Option.toOk("getCellAt(${ix})"))


    // def liftCellGetter(ix: Int32, 
    //                     fn: DataFormatter -> Cell -> a \ ef): RowReader[ka, ef, a] =
    //     liftGetterResultWithEnv((fmt, row) -> 
    //         match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
    //             case Some(cell) => fn(fmt, cell) |> Ok
    //             case None       => Err("liftCellGetter(${ix})")
    //         }
    //     )

    // def liftCellGetterResult(ix: Int32, 
    //                         fn: DataFormatter -> Cell -> Result[err, a] \ ef): RowReader[ka, ef, a] with ToString[err] =
    //     liftGetterResultWithEnv((fmt, row) -> 
    //         match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
    //             case Some(cell) => fn(fmt, cell) |> Result.mapErr(ToString.toString)
    //             case None       => Err("liftCellGetterResult(${ix})")
    //         }
    //     )

    pub def getRowNum(): RowReader[ka, ef, Int32] = 
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row)))


    pub def getSheetName(): RowReader[ka, ef, String] = 
        liftGetter(row -> {
            let sheet = Interop.Apache.Poi.SS.Usermodel.Row.getSheet(row);
            checked_ecast(Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName(sheet))
        })
            

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    /// Effect not sufficiently polymorphic...
    pub def getColumnIndex(s: String): RowReader[ka, ef, Int32] = 
        liftAction(_ -> 
            let addr = Interop.Apache.Poi.SS.Util.CellAddress.cellAddressWithAddress(s + "1");
            checked_ecast(Interop.Apache.Poi.SS.Util.CellAddress.getColumn(addr))
        )


    pub def getFirstCellNum(): RowReader[ka, ef, Int32] =
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getFirstCellNum(row)) |> Int16.toInt32)

    pub def getLastCellNum(): RowReader[ka, ef, Int32] =
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getLastCellNum(row)) |> Int16.toInt32)
            |> map(n -> if (n > 0) n - 1 else n)


/*
    /// Uses the default DataFormatter
    pub def getContentAt(ix: Int32): RowReader[ka, ef, String] =
        liftCellGetter(ix, (fmt, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, fmt))


    /// Uses the default DataFormatter
    pub def getTrimmedContentAt(ix: Int32): RowReader[ka, ef, String] =
        map(String.trim, getContentAt(ix))

    pub def isNullAt(ix: Int32): RowReader[ka, ef, Bool] = 
        getCellAt(ix) <$$> Object.isNull

    pub def getNullAt(ix: Int32): RowReader[ka, ef, Unit] = 
        getCellAt(ix)   >>= cell -> 
        if (Object.isNull(cell)) point() else throwError("getNullAt(${ix})")


    pub def getCellTypeAt(ix: Int32): RowReader[ka, ef, CellType] =
        liftCellGetter(ix, (_, cell) -> {
            Interop.Apache.Poi.SS.Usermodel.Cell.getCellType(cell) 
                |> SheetIO.CellType.fromNative
        })


    pub def isBlankAt(ix: Int32): RowReader[ka, ef, Bool] = 
        getCellTypeAt(ix) >>= typ ->
        match typ { 
            case TyBlank => point(true)
            case _       => point(false)
        } |> swapError("isBlankAt(${ix})")


    pub def getBlankAt(ix: Int32): RowReader[ka, ef, Unit] = 
        let errMsg = "getBlankAt(${ix})";
        getCellTypeAt(ix) >>= typ ->
        match typ { 
            case TyBlank => point()
            case _       => throwError(errMsg)
        } |> swapError(errMsg)

      



    /// List (not Array) avoids region allocation...
    pub def getRowContents(): RowReader[ka, ef, Vector[String]] =
        getRowContentsWith(getContentAt)


    pub def getRowContentsWith(evalCell: Int32 -> RowReader[ka, ef, a]): RowReader[ka, ef, Vector[a]] =
        getLastCellNum()                                            >>= end ->
        getRowContentsWithHelper(evalCell, end, 0, ks -> point(ks)) >>= xs  ->
        point(List.toVector(xs))


    // Should be tail recursive as bind is in the contiuation
    def getRowContentsWithHelper(evalCell: Int32 -> RowReader[ka, ef, a], 
                                    end: Int32, 
                                    ix: Int32, 
                                    k: List[a] -> RowReader[ka, ef, List[a]]): RowReader[ka, ef, List[a]] =
        if (ix > end)
            k(Nil)
        else 
            getRowContentsWithHelper(evalCell, end, ix + 1, ss -> evalCell(ix) >>= s -> k(s :: ss))

    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(): RowReader[ka, ef, Vector[String]]  =
        let eval = ix -> GradedOr.orelse(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(eval)

    def getIsoLocalDateTimeAsString(ix: Int32): RowReader[ka, ef, String] = 
        let fmt = Time.Format.DateTimeFormatter.iso_local_date_time();
        getLocalDateTimeAt(ix)                              >>= dt  ->
        point(Time.LocalDateTime.format(fmt, dt))

*/

}
