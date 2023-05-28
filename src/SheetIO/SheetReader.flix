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


    pub enum SheetReader[r: Region, ka: Type, ef: Eff, a: Type](
        Graded.Evaluator[ka, Interop.Apache.Poi.SS.Usermodel.Sheet[r], ef, a])


    instance GradedFunctor[SheetReader[r, ka]] {
        pub def map(f: a -> b \ ef1, 
                    ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1, b] = 
            SheetIO.SheetReader.map(f, ma)
    }

    instance GradedApplicative[SheetReader[r, ka]] {
        pub def point(x: a): SheetReader[r, ka, ef, a] = 
            SheetIO.SheetReader.point(x)

        pub def ap(mf: SheetReader[r, ka, ef1, a -> b \ ef2], 
                    ma: SheetReader[r, ka, ef, a]) : SheetReader[r, ka, ef + ef1 + ef2, b] = 
            SheetIO.SheetReader.ap(mf, ma)
    }

    instance GradedMonad[SheetReader[r, ka]] {
        pub def flatMap(f: a -> SheetReader[r, ka, ef1, b] \ ef2, 
                        ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1 + ef2, b] =
            SheetIO.SheetReader.flatMap(f, ma)
    }

    instance GradedZero[SheetReader[r, ka]] {
        pub def empty(): SheetReader[r, ka, ef, a] = 
            SheetIO.SheetReader.empty()
    }

    instance GradedOr[SheetReader[r, ka]] {
        pub def orelse(ma: SheetReader[r, ka, ef1, a], 
                        mb: SheetReader[r, ka, ef2, a]): SheetReader[r, ka, ef1 + ef2, a] =
            SheetIO.SheetReader.orelse(ma, mb)
    }

    instance GradedError[SheetReader[r, ka]] {
        pub def throwError(msg: String): SheetReader[r, ka, ef, a] = 
            SheetIO.SheetReader.throwError(msg)

        pub def mapError(f: String -> String \ ef1, 
                            ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1, a] =
            SheetIO.SheetReader.mapError(f, ma)

        pub def catchError(ma: SheetReader[r, ka, ef, a], 
                            f: String -> SheetReader[r, ka, ef1, a] \ ef2): SheetReader[r, ka, ef + ef1 + ef2, a] =
            SheetIO.SheetReader.catchError(ma, f)

        pub def reifyError(ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef, Result[String, a]] =
            SheetIO.SheetReader.reifyError(ma)
    }


}

mod SheetIO.SheetReader {

    use GradedMonad.{>>=}

    use SheetIO.SheetReader
    use SheetIO.SheetReader.{SheetReader}

    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.Row
    use Interop.Apache.Poi.SS.Usermodel.Sheet

    use SheetIO.RowEvaluator
    use SheetIO.NamedRowEvaluator
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator}



    pub def getSheetReader(x: SheetReader[r, ka, ef, a]): Graded.Evaluator[ka, Sheet[r], ef, a] = 
        let SheetReader(x1) = x;
        x1

    pub def runSheetReader(ma: SheetReader[r, a, ef, a],
                            s: Sheet[r]): Result[String, a] \ ef =
        Graded.Evaluator.runEvaluator(getSheetReader(ma), s)

    pub def point(x: a): SheetReader[r, ka, ef, a] = 
        SheetReader(Graded.Evaluator.point(x))

    pub def map(f: a -> b \ ef1, 
                ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1, b] =
        SheetReader(Graded.Evaluator.map(f, getSheetReader(ma)))

    pub def ap(mf: SheetReader[r, ka, ef1, a -> b \ ef2], 
                ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1 + ef2, b] =
        SheetReader(Graded.Evaluator.ap(getSheetReader(mf), getSheetReader(ma)))

    pub def flatMap(f: a -> SheetReader[r, ka, ef1, b] \ ef2, 
                    ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1 + ef2, b] =
        SheetReader(Graded.Evaluator.flatMap(f >> getSheetReader, getSheetReader(ma)))

    pub def throwError(msg: String): SheetReader[r, ka, ef, a] = 
        SheetReader(Graded.Evaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef1, 
                        ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef + ef1, a] =
        SheetReader(Graded.Evaluator.mapError(f, getSheetReader(ma)))

    pub def catchError(ma: SheetReader[r, ka, ef, a], 
                        f: String -> SheetReader[r, ka, ef1, a] \ ef2): SheetReader[r, ka, ef + ef1 + ef2, a] =
        SheetReader(Graded.Evaluator.catchError(getSheetReader(ma), f >> getSheetReader))

    pub def reifyError(ma: SheetReader[r, ka, ef, a]): SheetReader[r, ka, ef, Result[String, a]] =
        SheetReader(Graded.Evaluator.reifyError(getSheetReader(ma)))

    pub def empty(): SheetReader[r, ka, ef, a] = 
        SheetReader(Graded.Evaluator.empty())
        
    pub def orelse(ma: SheetReader[r, ka, ef1, a], 
                    mb: SheetReader[r, ka, ef2, a]): SheetReader[r, ka, ef1 + ef2, a] =
        SheetReader(Graded.Evaluator.orelse(getSheetReader(ma), getSheetReader(mb)))

    def liftGetter(f: Sheet[r] -> a \ ef): SheetReader[r, ka, ef, a] =
        SheetReader(Graded.Evaluator.liftGetter(f))
    
    def liftGetterResult(f: Sheet[r] -> Result[err, a] \ ef): SheetReader[r, ka, ef, a] with ToString[err] =
        SheetReader(Graded.Evaluator.liftGetterResult(f))

/*

    pub def liftAction(f: Unit -> a \ ef): SheetReader[r, ka, a] \ ef =
        SheetReader(Graded.Evaluator.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[err, a] \ ef): SheetReader[r, ka, a] \ ef with ToString[err] =
        SheetReader(Graded.Evaluator.liftActionResult(f))


    // pub def testM(ma: SheetReader[r, ka, a]): SheetReader[r, ka, Bool] =
    //     SheetReader(wb -> 
    //         match apply1(ma, wb) { 
    //             case Ok(_)  => Ok(true)
    //             case Err(_) => Ok(false)
    //         }
    //     )

    ///
    /// Expose the result of operation `ma`.
    ///
    /// By definition this never fails.
    ///
    // pub def result(ma: SheetReader[r, ka, a]): SheetReader[r, ka, Result[a, String]] = 
    //     SheetReader(wb -> 
    //         let ans = apply1(ma, wb) |> Result.mapErr(SheetIO.Internal.EvalError.getText);
    //         Ok(ans)
    //     )



*/


    // SheetReader
 

    pub def getSheetName(): SheetReader[r, ka, r, String] =
        liftGetter(Interop.Apache.Poi.SS.Usermodel.Sheet.getSheetName)


    pub def getFirstRowNum(): SheetReader[r, ka, r, Int32] =
        liftGetterResult(Interop.Apache.Poi.SS.Usermodel.Sheet.getFirstRowNum >> Option.toOk("getFirstRowNum"))

    /// This may be higher than expected if rows at the end of the sheet have previously had content 
    /// which has been beleted.
    pub def getLastRowNum(): SheetReader[r, ka, r, Int32] =
        liftGetterResult(Interop.Apache.Poi.SS.Usermodel.Sheet.getLastRowNum >> Option.toOk("getLastRowNum"))


    pub def getPhysicalNumberOfRows(): SheetReader[r, ka, r, Int32] =
        liftGetter(Interop.Apache.Poi.SS.Usermodel.Sheet.getPhysicalNumberOfRows)


    pub def getRowAt(ix: Int32, eval: RowEvaluator[a, r, ef, a]): SheetReader[r, ka, ef + r, a] =
        use Result.flatMap;
        liftGetterResult(sheet -> {
            /// TODO - df should be "user facing"...
            forM (df <- region rc {Interop.Apache.Poi.SS.Usermodel.DataFormatter.new(rc, _ -> checked_ecast(Ok()))};
                    row <- Interop.Apache.Poi.SS.Usermodel.Sheet.getRow(ix, sheet);
                    ans <- SheetIO.RowEvaluator.runRowEvaluator(eval, df, row))
                yield (ans)
        })
        

    ///
    /// Working with Iterator<Row> - wrapped here as a Stream - seems safer from errors 
    /// than trying to use `getLastRowNum` to find the end.
    ///
    // def getRowIterator(rc: Region[r1]): SheetReader[r, ka, r, Iterator[Row[r1], r1, r1]] =
    //     liftGetter(Interop.Apache.Poi.SS.Usermodel.Sheet.rowIterator(rc))


    pub def getHeaders(headerIndex: Int32): SheetReader[r, ka, r, Vector[String]] =
        getRowAt(headerIndex, SheetIO.RowEvaluator.getRowContents())


/*

    /// This might be expressive enough to write all `RowEvaluator` traversals with...
    def foldRowsTraversal(range: {start = Int32, end = Int32}, 
                            skipErrors: Bool,
                            f: (b, a, Int32) -> b \ ef, 
                            s: b, 
                            eval: RowEvaluator[r, a, a]): SheetReader[r, ka, b] \ { r, ef } = 
        let* rowIter = getRowIterator();
        def loop(s1) = match Iterator.next(rowIter) {
            case None      => point(s1)
            case Some(row) => {
                let r1 = Scoped.regionOf(row);
                let rownum = Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row);
                if (rownum >= range.start and rownum < range.end) 
                    match SheetIO.RowEvaluator.runRowEvaluator(r1, eval, row) {
                        case Err(_)  => if (not skipErrors) throwError("foldRowsTraversal") else loop(s1)
                        case Ok(ans) => loop(f(s1, ans, rownum))
                    }
                else 
                    loop(s1)
            }
        };
        loop(s)


    /// This might be expressive enough to write all `RowEvaluator` traversals with...
    def foldNamedRowsTraversal(range: {start = Int32, end = Int32}, 
                                skipErrors: Bool,
                                headers: List[String],
                                f: (b, a, Int32) -> b \ ef, 
                                s: b, 
                                eval: NamedRowEvaluator[r, a, a]): SheetReader[r, ka, b] \ { r, ef } =
        let* rowIter        = getRowIterator();
        def loop(s1) = match Iterator.next(rowIter) {
            case None     => point(s1)
            case Some(row) => {
                let r1     = Scoped.regionOf(row);
                let rownum = Interop.Apache.Poi.SS.Usermodel.Row.getRowNum(row);
                if (rownum >= range.start and rownum < range.end) 
                    match SheetIO.NamedRowEvaluator.runNamedRowEvaluator(r1, eval, headers, row) {
                        case Err(_)  => if (not skipErrors) throwError("foldNamedRowsTraversal") else loop(s1)
                        case Ok(ans) => loop(f(s1, ans, rownum))
                    }
                else 
                    loop(s1)
            }
        };
        loop(s)

    pub def allRows(eval: RowEvaluator[r, a, a]): SheetReader[r, ka, List[a]] \ { r } =
        let* end = getPhysicalNumberOfRows();
        foldRowsTraversal({start = 0, end = end}, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))


    /// rows where `eval` fails are discounted...
    pub def someRows(eval: RowEvaluator[r, a, a]): SheetReader[r, ka, List[a]] \ { r } =
        let* end = getPhysicalNumberOfRows();
        foldRowsTraversal({start = 0, end = end}, true, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    /// WARNING - start will be "miscalculated" if there are blank rows between 0 and start.
    /// We need a more robust implementation than this...

    pub def allRowsFrom(start: Int32, eval: RowEvaluator[r, a, a]): SheetReader[r, ka, List[a]] \ { r } =
        let* end = getPhysicalNumberOfRows();
        foldRowsTraversal({start = start, end = end}, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    pub def rowsInRange(range: {start = Int32, end = Int32}, eval: RowEvaluator[r, a, a]): SheetReader[r, ka, List[a]] \ { r } =
        foldRowsTraversal(range, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))


    /// Headers at 0, start at 1
    pub def allNamedRows(eval: NamedRowEvaluator[r, a, a]): SheetReader[r, ka, List[a]] \ { r } =
        let* headers = getHeaders(0);
        let* end = getPhysicalNumberOfRows();
        foldNamedRowsTraversal({start = 1, end = end}, false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    pub def allNamedRowsFrom(pos: {headerIndex = Int32, start = Int32}, eval: NamedRowEvaluator[r, a, a]): SheetReader[r, ka, List[a]] \ { r } =
        let* headers = getHeaders(pos.headerIndex);
        let* end = getPhysicalNumberOfRows();
        foldNamedRowsTraversal({start = pos.start, end = end}, false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))


    /// Row 0 is read for the headers, some subsequent rows are *data*.
    // pub def someRows(eval: RowEvaluator[r, Option[a]], sheet: Sheet): SheetReader[r, ka, List[a]] = 
    //     let postprocess = xs -> List.foldRightWithCont((x, fac) -> match x {
    //             case None => fac()
    //             case Some(x1) => x1 :: fac()
    //         }, Nil, xs);
    //     map(postprocess, allRows(eval, sheet))


    

    // pub def rowEval(fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
    //                 eval: RowEvaluator[r, a], 
    //                 x: Row): SheetReader[r, ka, a] =
    //     let RowEvaluator(f) = eval;
    //     let Row(x1) = x;
    //     liftActionResult(_ -> f(fmt, x1) |> Result.mapErr(SheetIO.EvalError.toString))



    // pub def namedRowsInRange(ixs: Map[String, Int32], b: Int32, e: Int32, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, ka, List[a]] = 
    //     let NamedRowEvaluator(f) = eval;
    //     let checker         = row -> {
    //         let ix = SheetIO.Internal.Row.getRowNum(row);
    //         if (ix >= b and ix < e) Some(row) else None
    //     };
    //     let* fmt            = liftAction(_ -> SheetIO.Internal.DataFormatter.new());
    //     let* strm           = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox.Deprecated.Cursor.filterMap(checker));
    //     let* xs             = liftActionResult(_ -> FlixSandbox.Deprecated.Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO.EvalError.toString), strm) |> FlixSandbox.Deprecated.Cursor.toList!);
    //     point(xs)

    // pub def nameRowEval(indices: Map[String, Int32], 
    //                     fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
    //                     eval: NamedRowEvaluator[r, a], x: Row): SheetReader[r, ka, a] =
    //     let NamedRowEvaluator(f) = eval;
    //     let Row(x1) = x;
    //     liftActionResult(_ -> f(indices, fmt, x1) |> Result.mapErr(SheetIO.EvalError.toString))



    /// Reader headers from row at `headerIndex` then read data rows from `start`.
    /// Note - might be wise to change back to flix-sandbox Cursor while stdlib Interator
    /// is in development.
    /// `headerIndex` and `start` are zero indexed (as per POI rather than Excel).
    ///
    // pub def namedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, ka, List[a]] =   
    //     let NamedRowEvaluator(f)    = eval;
    //     let checker         = row -> {
    //         let ix = SheetIO.Internal.Row.getRowNum(row);
    //         if (ix >= r.start) Some(row) else None
    //     };
    //     let* fmt        = liftAction(_ -> SheetIO.Internal.DataFormatter.new());
    //     let* row0       = getRowAt(sheet, r.headerIndex);
    //     let* ixs        = region r1 {
    //         let* hdrs = rowEval(fmt, SheetIO.RowEvaluator.getRowContents(r1), row0);
    //         point(SheetIO.NamedRowEvaluator.makeHeaders(hdrs))
    //     };
    //     let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox.Deprecated.Cursor.filterMap(checker));
    //     let* xs         = liftActionResult(_ -> FlixSandbox.Deprecated.Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO.EvalError.toString), strm) |> FlixSandbox.Deprecated.Cursor.toList!);
    //     point(xs)


    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allNamedRows(eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, ka, List[a]] = 
        namedRowsFrom({headerIndex = 0, start = 1}, eval, sheet)

    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someNamedRows(eval: NamedRowEvaluator[r, Option[a]], sheet: Sheet): SheetReader[r, ka, List[a]] = 
        let postprocess = xs -> List.foldRightWithCont((x, fac) -> match x {
                case None => fac()
                case Some(x1) => x1 :: fac()
            }, Nil, xs);
        map(postprocess, allNamedRows(eval, sheet))


    pub def foldNamedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, 
                                combine: (b, a) -> b \ ef, ac: b, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, ka, b] =
        let NamedRowEvaluator(f) = eval;
        let checker         = row -> {
            let ix = SheetIO.Internal.Row.getRowNum(row);
            if (ix >= r.start) Some(row) else None
        };
        let* fmt        = liftAction(_ -> SheetIO.Internal.DataFormatter.new());
        let* row0       = getRowAt(sheet, r.headerIndex);
        let* ixs        = region r1 { 
            let* hdrs = rowEval(fmt, SheetIO.RowEvaluator.getRowContents(r1), row0);
            point(SheetIO.NamedRowEvaluator.makeHeaders(hdrs))
        };
        let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox.Deprecated.Cursor.filterMap(checker));
        let* ans        = liftActionResult(_ -> FlixSandbox.Deprecated.Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO.EvalError.toString), strm) |> FlixSandbox.Deprecated.Cursor.fold!(combine, ac));
        point(ans)


    pub def foldAllNamedRows(combine: (b, a) -> b \ ef, ac: b, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, ka, b] =
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
