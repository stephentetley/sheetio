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
    use Graded.EnvEvaluator.{ask, liftGetter, liftGetterResult, liftGetterResultWithEnv, liftAction, liftActionResult}
    use Time.Format.DateTimeFormatter
    

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.DataFormatter
    use Interop.Apache.Poi.SS.Usermodel.Row
    
    use SheetIO.RowReader
    use SheetIO.CellReader
    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    


    pub def runRowReader(ma: RowReader[a, ef, a], 
                            env: DataFormatter,
                            st: Row): Result[String, a] \ ef =
        Graded.EnvEvaluator.runEnvEvaluator(ma, env, st)
        

    pub def withCellReader(eval: CellReader[a, ef, a], cell: Cell): RowReader[ka, ef, a] =
        ask()       >>= df -> 
        liftActionResult(_ -> SheetIO.CellReader.runCellReader(eval, df, cell))


    pub def getCellAt(ix: Int32): RowReader[ka, ef, Cell] = 
        liftGetterResult(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row)) |> Option.toOk("getCellAt(${ix})"))

    pub def readCellAt(ix: Int32, eval: CellReader[a, ef, a]): RowReader[ka, ef, a] =
        getCellAt(ix)    >>= cell -> 
        withCellReader(eval, cell)

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



    pub def cellIterator(rc: Region[r]): RowReader[ka, ef + r, Iterator[Cell, r, r]] = 
        liftGetter(row -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Row.cellIterator(rc, row)))


    
    /// Ideally we would not need to expose `rc` but due to types within the Iterator
    /// it would otherwise leak...
    pub def getRowContents(rc: Region[r]): RowReader[ka, ef + r, Vector[String]] =
        getRowContentsWith(rc, SheetIO.CellReader.getContent())


    /// Ideally we would not need to expose `rc` but due to types within the Iterator
    /// it would otherwise leak...
    pub def getRowContentsWith(rc: Region[r], eval: CellReader[a, ef, a]): RowReader[ka, ef + r, Vector[a]] =
        let step = (acc, cell) -> { 
            withCellReader(eval, cell) >>= a ->
            point(Chain.snoc(acc, a))
        };
        cellIterator(rc) >>= iter -> 
        GradedExtras.Iterator.foldLeftM(step, Chain.empty(), iter) >>= chain -> 
        point(List.toVector(Chain.toList(chain)))

/*


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
