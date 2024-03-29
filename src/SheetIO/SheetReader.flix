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


    pub type alias SheetReader[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.RgnEnvEvaluator[ka, r, SheetIO.DataFormatter, SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet, ef, a]


}

mod SheetIO.SheetReader {

    use GradedFunctor.{<$$>}
    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{throwError, swapError}
    use Graded.RgnEnvEvaluator.{liftGetter, liftGetterResult, liftGetterWithRgn, liftActionResult}

    use SheetIO.SheetReader

    use SheetIO.Internal.Apache.Poi.SS.Usermodel.CellStyle
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Row
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet

    use SheetIO.RowReader
    use SheetIO.NamedRowReader




    pub def runSheetReader(ma: SheetReader[a, r, ef, a],
                            rc: Region[r], 
                            df: SheetIO.DataFormatter,
                            s: Sheet): Result[String, a] \ ef =
        Graded.RgnEnvEvaluator.runRgnEnvEvaluator(ma, rc, df, s)


    pub def withRowReader(eval: RowReader[a, r, ef, a], row: Row): SheetReader[ka, r, ef, a] =
        GradedExtras.Utils.liftGetterResultWithAll((rc, df, _) -> SheetIO.RowReader.runRowReader(eval, rc, df, row))

        
    pub def withNamedRowReader(eval: NamedRowReader[a, r, ef, a], headers: Vector[String], row: Row): SheetReader[ka, r, ef, a] =
        GradedExtras.Utils.liftGetterResultWithAll((rc, df, _) -> SheetIO.NamedRowReader.runNamedRowReader(eval, rc, headers, df, row))

    pub def getRowAt(ix: Int32): SheetReader[ka, r, ef, Row] = 
        liftGetterResult(sheet -> checked_ecast(SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.getRow(ix, sheet)))
    
    pub def readRowAt(ix: Int32, eval: RowReader[a, r, ef, a]): SheetReader[ka, r, ef, a] =
        getRowAt(ix)    >>= row -> 
        withRowReader(eval, row)
        
 

    pub def getSheetName(): SheetReader[ka, r, ef, String] =
        liftGetter(s -> checked_ecast(SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.getSheetName(s)))


    pub def getFirstRowNum(): SheetReader[ka, r, ef, Int32] =
        liftGetterResult(s -> checked_ecast(SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.getFirstRowNum(s)) |> Option.toOk("getFirstRowNum"))

    /// This may be higher than expected if rows at the end of the sheet have previously had content 
    /// which has been beleted.
    pub def getLastRowNum(): SheetReader[ka, r, ef, Int32] =
        liftGetterResult(s -> checked_ecast(SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.getLastRowNum(s)) |> Option.toOk("getLastRowNum"))


    pub def getPhysicalNumberOfRows(): SheetReader[ka, r, ef, Int32] =
        liftGetter(s -> checked_ecast(SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.getPhysicalNumberOfRows(s)))

    ///
    /// Working with Iterator<Row> - wrapped here as a Stream - seems safer from errors 
    /// than trying to use `getLastRowNum` to find the end.
    ///
    pub def rowIterator(): SheetReader[ka, r, ef + r, Iterator[Row, r, r]] = 
        liftGetterWithRgn((rc, s) -> checked_ecast(SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.rowIterator(rc, s)))



    pub def getHeaders(rowIx: Int32): SheetReader[ka, r, ef + r, Vector[String]] =
        getRowAt(rowIx) >>= row -> 
        withRowReader(SheetIO.RowReader.getRowContents(), row)
    



    /// This might be expressive enough to write all `RowReader` traversals with...
    def foldRowsTraversal(start: {start = Int32},
                            end: {end = Int32}, 
                            skipErrors: Bool,
                            f: (b, a, Int32) -> b \ ef1, 
                            s: b, 
                            eval: RowReader[a, r, ef, a]): SheetReader[ka, r, ef + ef1 + r, b] = 
        rowIterator() >>= iter -> 
        def loop(s1) = match Iterator.next(iter) {
            case None      => point(s1)
            case Some(row) => {
                let rownum = SheetIO.Internal.Apache.Poi.SS.Usermodel.Row.getRowNum(row);
                if (rownum >= start.start and rownum < end.end) {
                    GradedError.reifyError(withRowReader(eval, row)) >>= ans ->
                    match ans {
                        case Err(_) => if (not skipErrors) throwError("foldRowsTraversal") else loop(s1)
                        case Ok(a1) => loop(f(s1, a1, rownum))
                    }
                } else loop(s1)
            }
        };
        loop(s)


    /// This might be expressive enough to write all `RowReader` traversals with...
    def foldNamedRowsTraversal(start: {start = Int32}, 
                                end: {end = Int32}, 
                                skipErrors: Bool,
                                headers: Vector[String],
                                f: (b, a, Int32) -> b \ ef1, 
                                s: b, 
                                eval: NamedRowReader[a, r, ef, a]): SheetReader[ka, r, ef + ef1 + r, b] =
       rowIterator() >>= iter -> 
        def loop(s1) = match Iterator.next(iter) {
            case None      => point(s1)
            case Some(row) => {
                let rownum = SheetIO.Internal.Apache.Poi.SS.Usermodel.Row.getRowNum(row);
                if (rownum >= start.start and rownum < end.end) {
                    GradedError.reifyError(withNamedRowReader(eval, headers, row)) >>= ans ->
                    match ans {
                        case Err(_) => if (not skipErrors) throwError("foldNamedRowsTraversal") else loop(s1)
                        case Ok(a1) => loop(f(s1, a1, rownum))
                    }
                } else loop(s1)
            }
        };
        loop(s)


    pub def allRows(eval: RowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        foldRowsTraversal(start = 0, end = Int32.maxValue(), false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))

    /// rows where `eval` fails are discounted...
    pub def someRows(eval: RowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        foldRowsTraversal(start = 0, end = Int32.maxValue(), true, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))

    /// WARNING - start will be "miscalculated" if there are blank rows between 0 and start.
    /// We need a more robust implementation than this...

    pub def allRowsFrom(start: Int32, eval: RowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        foldRowsTraversal(start = start, end = Int32.maxValue(), false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))

    pub def rowsInRange(start: {start = Int32}, end: {end = Int32}, eval: RowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        foldRowsTraversal(start, end, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))

    /// Headers at 0, start at 1
    pub def allNamedRows(eval: NamedRowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        getHeaders(0) >>= headers ->
        foldNamedRowsTraversal(start = 1, end = Int32.maxValue(), false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))

    /// rows where `eval` fails are discounted...
    pub def someNamedRows(eval: NamedRowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        getHeaders(0) >>= headers ->
        foldNamedRowsTraversal(start = 1, end = Int32.maxValue(), true, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))


    pub def allNamedRowsFrom(headerIndex: {headerIndex = Int32}, 
                                start: {start = Int32}, 
                                eval: NamedRowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        getHeaders(headerIndex.headerIndex) >>= headers ->
        foldNamedRowsTraversal(start, end = Int32.maxValue(), false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))


    pub def namedRowsInRange(headerIndex: {headerIndex = Int32}, 
                                start: {start = Int32}, 
                                end: {end = Int32}, 
                                eval: NamedRowReader[a, r, ef, a]): SheetReader[ka, r, ef + r, List[a]] =
        getHeaders(headerIndex.headerIndex) >>= headers ->
        foldNamedRowsTraversal(start, end, false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) 
            |> GradedFunctor.map(k -> k(Nil))


/*





    pub def foldNamedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, 
                                combine: (b, a) -> b \ ef, ac: b, eval: NamedRowReader[r, a], sheet: Sheet): SheetReader[r, ka, b] =
        let NamedRowReader(f) = eval;
        let checker         = row -> {
            let ix = SheetIO.Internal.Row.getRowNum(row);
            if (ix >= r.start) Some(row) else None
        };
        let* fmt        = liftAction(_ -> SheetIO.Internal.DataFormatter.new());
        let* row0       = getRowAt(sheet, r.headerIndex);
        let* ixs        = region r1 { 
            let* hdrs = rowEval(fmt, SheetIO.RowReader.getRowContents(r1), row0);
            point(SheetIO.NamedRowReader.makeHeaders(hdrs))
        };
        let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox.Deprecated.Cursor.filterMap(checker));
        let* ans        = liftActionResult(_ -> FlixSandbox.Deprecated.Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO.EvalError.toString), strm) |> FlixSandbox.Deprecated.Cursor.fold!(combine, ac));
        point(ans)


    pub def foldAllNamedRows(combine: (b, a) -> b \ ef, ac: b, eval: NamedRowReader[r, a], sheet: Sheet): SheetReader[r, ka, b] =
        foldNamedRowsFrom({headerIndex = 0, start = 1}, combine, ac, eval, sheet)


    ///
    /// Working with Iterator<Row> - wrapped here as a Stream - seems safer from errors 
    /// than trying to use `getLastRowNum` to find the end.
    ///
    def getRowCursor(x: Sheet): Cursor[String, ##org.apache.poi.ss.usermodel.Row] \ IO =
        import new flixspt.sheetio.POIRowIterator(##org.apache.poi.ss.usermodel.Sheet): ##flixspt.sheetio.POIRowIterator \ IO as newRowCursor;
        import flixspt.sheetio.POIRowIterator.hasNext(): Bool \ IO as ffiHasNext;
        import flixspt.sheetio.POIRowIterator.next(): ##org.apache.poi.ss.usermodel.Row \ IO as ffiNext;
        try {
            let Sheet(x1)       = x;
            let iter = newRowCursor(x1);
            let cursorAlg =  {
                  hasNext   = () -> ffiHasNext(iter)
                , next      = FlixSandbox.Deprecated.Cursor.guardThunk(_ -> ffiNext(iter), "next failed")
                , finalize  = () -> () as \ IO
            };
            FlixSandbox.Deprecated.Cursor.makeCursor(cursorAlg)
        } catch {
            case _: ##java.lang.Exception => FlixSandbox.Deprecated.Cursor.failure("newPOIRowStream")
        }

    pub def mapWithIndexM(mf: (a, Int32) -> SheetReader[r, ka, b], xs: List[a]): SheetReader[r, ka, List[b]] =
        SheetReader(wb ->
            mapWithIndexMHelper(wb, mf, 0, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            mf: (a, Int32) -> SheetReader[r, ka, b], 
                            ix : Int32, 
                            xs: List[a], 
                            fk: String -> Result[List[b], String], 
                            sk: List[b] -> Result[List[b], String]): Result[List[b], String] \ IO =
        match xs {
            case Nil => sk(Nil)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb) {
                    case Err(e) => fk(e)
                    case Ok(x1) => mapWithIndexMHelper(wb, mf, ix + 1, rs, fk, ks -> sk(x1 :: ks))
                }
            }
        }

    pub def mapWithIndexX(mf: (a, Int32) -> SheetReader[r, ka, b], xs: List[a]): SheetReader[r, ka, Unit] =
        SheetReader(wb ->
            mapWithIndexXHelper(wb, mf, 0, xs, e -> Err(e), _ -> Ok())
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            mf: (a, Int32) -> SheetReader[r, ka, b],
                            ix: Int32, 
                            xs: List[a], 
                            fk: String -> Result[Unit, String], 
                            sk: Unit -> Result[Unit, String]): Result[Unit, String] \ IO =
        match xs {
            case Nil => sk()
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb) {
                    case Err(e) => fk(e)
                    case Ok(_) => mapWithIndexXHelper(wb, mf, ix + 1, rs, fk, sk)
                }
            }
        }


    ///
    /// Returns `(v1 :: v2 :: ... v :: vn)` if each of `f(xs_i)` is successful.
    /// Otherwise fails with the first error encountered.
    ///
    /// `traverse` is a synonym for `mapM`.
    ///
    pub def traverse(f: a -> SheetReader[r, ka, b], xs: List[a]): SheetReader[r, ka, List[b]] = mapM(f, xs)

    ///
    /// `traverseX` is a synonym for `mapX`.
    ///
    pub def traverseX(f: a -> SheetReader[r, ka, b], xs: List[a]): SheetReader[r, ka, Unit] = mapX(f, xs)

    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    pub def forM(xs: List[a], f: a -> SheetReader[r, ka, b]): SheetReader[r, ka, List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    pub def forX(xs: List[a], f: a -> SheetReader[r, ka, b]): SheetReader[r, ka, Unit] = mapX(f, xs)


    pub def foldM(mf: (b, a) -> SheetReader[r, ka, b], start: b, xs: List[a]): SheetReader[r, ka, b] =
        SheetReader(wb ->
            foldMHelper(wb, mf, start, xs, e -> Err(e), ac -> Ok(ac))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    mf: (b, a) -> SheetReader[r, ka, b], 
                    acc: b, 
                    xs: List[a], 
                    fk: String -> Result[b, String], 
                    sk: b -> Result[b, String]): Result[b, String] \ IO =
        match xs {
            case Nil => sk(acc)
            case x :: rs => {
                let ma = mf(acc, x);
                match apply1(ma, wb) {
                    case Err(e) => fk(e)
                    case Ok(acc1) => foldMHelper(wb, mf, acc1, rs, fk, sk)
                }
            }
        }

    pub def mapAccumM(mf: (ac, a) -> SheetReader[r, ka, (ac, b)], start: ac, xs: List[a]): SheetReader[r, ka, (ac, List[b])] =
        SheetReader(wb ->
            mapAccumMHelper(wb, mf, start, xs, e -> Err(e), (ac, ks) -> Ok((ac, ks)))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        mf: (ac, a) -> SheetReader[r, ka, (ac, b)],
                        ac: ac,
                        xs: List[a],
                        fk: String -> Result[(ac, List[b]), String],
                        sk: (ac, List[b]) -> Result[(ac, List[b]), String]): Result[(ac, List[b]), String] \ IO =
        match xs {
            case Nil => sk(ac, Nil)
            case x :: rs => {
                let ma = mf(ac, x);
                match apply1(ma, wb) {
                    case Err(e) => fk(e)
                    case Ok((ac1, x1)) => mapAccumMHelper(wb, mf, ac1, rs, fk, (ac2, ks) -> sk(ac2, x1 :: ks))
                }
            }
        }

    ///
    /// `mapAccumX` is a synonym for `foldM`.
    ///
    pub def mapAccumX(mf: (ac, a) -> SheetReader[r, ka, ac], start: ac, xs: List[a]): SheetReader[r, ka, ac] = foldM(mf, start, xs)


*/

}
