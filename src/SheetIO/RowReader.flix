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

    pub type alias RowReader[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.RgnEnvEvaluator[ka, r, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Row, ef, a]

}

mod SheetIO.RowReader {

    
    use GradedFunctor.{map, <$$>}
    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{throwError, swapError}
    use Graded.RgnEnvEvaluator.{ask, liftGetter, liftGetterResult, liftGetterResultWithEnv, liftAction, liftActionResult}
    use Time.Format.DateTimeFormatter
    

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter
    use Interop.Apache.Poi.SS.Usermodel.Row
    
    use SheetIO.RowReader
    use SheetIO.CellReader
    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    


    pub def runRowReader(ma: RowReader[a, r, ef, a], 
                            rc: Region[r],
                            env: DataFormatter,
                            st: Row): Result[String, a] \ ef =
        Graded.RgnEnvEvaluator.runRgnEnvEvaluator(ma, rc, env, st)
        

    pub def withCellReader(eval: CellReader[a, ef, a], cell: Cell): RowReader[ka, r, ef, a] =
        ask()       >>= df -> 
        liftActionResult(_ -> SheetIO.CellReader.runCellReader(eval, df, cell))


    pub def getCellAt(ix: Int32): RowReader[ka, r, ef, Cell] = 
        liftGetterResult(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row)) |> Option.toOk("getCellAt(${ix})"))

    pub def readCellAt(ix: Int32, eval: CellReader[a, ef, a]): RowReader[ka, r, ef, a] =
        getCellAt(ix)    >>= cell -> 
        withCellReader(eval, cell)

    pub def getRowNum(): RowReader[ka, r, ef, Int32] = 
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row)))


    pub def getSheetName(): RowReader[ka, r, ef, String] = 
        liftGetter(row -> {
            let sheet = Interop.Apache.Poi.SS.Usermodel.Row.getSheet(row);
            checked_ecast(Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName(sheet))
        })
            

    /// Get the column index from the column identifier "A", "B", ... "AA", "AB", ...
    /// Effect not sufficiently polymorphic...
    pub def getColumnIndex(s: String): RowReader[ka, r, ef, Int32] = 
        liftAction(_ -> 
            let addr = Interop.Apache.Poi.SS.Util.CellAddress.cellAddressWithAddress(s + "1");
            checked_ecast(Interop.Apache.Poi.SS.Util.CellAddress.getColumn(addr))
        )


    pub def getFirstCellNum(): RowReader[ka, r, ef, Int32] =
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getFirstCellNum(row)) |> Int16.toInt32)

    pub def getLastCellNum(): RowReader[ka, r, ef, Int32] =
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getLastCellNum(row)) |> Int16.toInt32)
            |> map(n -> if (n > 0) n - 1 else n)



    // pub def cellIterator(rc: Region[r]): RowReader[ka, r, ef + r, Iterator[Cell, r, r]] = 
    //     liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.cellIterator(rc, row)))

    

    pub def cellIterator(): RowReader[ka, r, ef + r, Iterator[Cell, r, r]] = 
        Graded.RgnEnvEvaluator.liftGetterWithRgn((rc, row) -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.cellIterator(rc, row)))


    
    /// Ideally we would not need to expose `rc` but due to types within the Iterator
    /// it would otherwise leak...
    pub def getRowContents(): RowReader[ka, r, ef + r, Vector[String]] = 
        getRowContentsWith(SheetIO.CellReader.getContent())


    // /// Ideally we would not need to expose `rc` but due to types within the Iterator
    // /// it would otherwise leak...
    // pub def getRowContentsWith(rc: Region[r], eval: CellReader[a, ef, a]): RowReader[ka, r, ef + r, Vector[a]] =
    //     let step = (acc, cell) -> { 
    //         withCellReader(eval, cell) >>= a ->
    //         point(Chain.snoc(acc, a))
    //     };
    //     cellIterator(rc) >>= iter -> 
    //     GradedIterator.foldLeftM(step, Chain.empty(), iter) >>= chain -> 
    //     point(List.toVector(Chain.toList(chain)))


    /// Experiment - lift into a monad with region capabilities...
    pub def getRowContentsWith(eval: CellReader[a, ef, a]): RowReader[ka, r, ef + r, Vector[a]] =
        let step = (acc, cell) -> { 
            withCellReader(eval, cell) >>= a ->
            point(Chain.snoc(acc, a))
        };
        cellIterator() >>= iter -> 
        GradedIterator.foldLeftM(step, Chain.empty(), iter) >>= chain -> 
        point(List.toVector(Chain.toList(chain)))



    // def getIteratorCellContents(eval: CellReader[a, ef, a], iter: Iterator[Cell, r, r]): RowReader[ka, r, ef + r, Vector[a]] \ r =
    //     let step = (acc, cell) -> { 
    //         withCellReader(eval, cell) >>= a ->
    //         point(Chain.snoc(acc, a))
    //     }; 
    //     GradedExtras.Iterator.foldLeftM(step, Chain.empty(), iter) >>= chain -> 
    //     point(List.toVector(Chain.toList(chain)))

/*


    /// Special version of `getRowContents` - date-times are rewritten to ISO local date time strings.
    /// This should avoid date mangling.
    pub def getRowContentsIsoLocalDateTime(): RowReader[ka, r, ef, Vector[String]]  =
        let eval = ix -> GradedOr.orelse(getIsoLocalDateTimeAsString(ix), getContentAt(ix));
        getRowContentsWith(eval)

    def getIsoLocalDateTimeAsString(ix: Int32): RowReader[ka, r, ef, String] = 
        let fmt = Time.Format.DateTimeFormatter.iso_local_date_time();
        getLocalDateTimeAt(ix)                              >>= dt  ->
        point(Time.LocalDateTime.format(fmt, dt))

*/

}
