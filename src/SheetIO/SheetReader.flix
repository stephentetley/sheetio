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

    use SheetIO/Internal/EvalError.{EvalError};
    
    pub enum SheetReader[r: Region, a: Type](
        ##org.apache.poi.ss.usermodel.Sheet -> Result[a, EvalError] & r
    )

    instance Functor[SheetReader[r]] {
        pub def map(f: a -> b & ef, o: SheetReader[r, a]): SheetReader[r, b] & ef = 
            SheetIO/SheetReader.map(f, o)
    }

    instance Applicative[SheetReader[r]] {
        pub def point(x: a): SheetReader[r, a] = 
            SheetIO/SheetReader.point(x)

        pub def ap(f: SheetReader[r, a -> b & ef], x: SheetReader[r, a]) : SheetReader[r, b] & ef = 
            SheetIO/SheetReader.ap(f, x)
    }

    instance Monad[SheetReader[r]] {
        pub def flatMap(f: a -> SheetReader[r, b] & ef, o: SheetReader[r, a]): SheetReader[r, b] & ef =
            SheetIO/SheetReader.flatMap(f, o)
    }

    instance MonadZero[SheetReader[r]] {
        pub def empty(): SheetReader[r, a] = SheetIO/SheetReader.throwError("empty")
    }

    instance MonadLib.MonadOr[SheetReader[r]] {
        pub def orelse(a: SheetReader[r, a], b: SheetReader[r, a]): SheetReader[r, a] =
            SheetIO/SheetReader.orelse(a, b)
    }

}

namespace SheetIO/SheetReader {

    use Monad.{>>=};


    use SheetIO.SheetReader;
    use SheetIO.SheetReader.{SheetReader};

    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO/Internal/PrimError.{PrimError};

    use SheetIO.Row;
    use SheetIO.Row.{Row};
    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.Sheet;
    use SheetIO.Sheet.{Sheet};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};


    def apply1(ma: SheetReader[r, a], primSheet: ##org.apache.poi.ss.usermodel.Sheet): Result[a, EvalError] & r = 
        let SheetReader(x1) = ma;
        x1(primSheet)

    pub def evaluate(ma: SheetReader[r, a], sheet: Sheet[r]): Result[a, EvalError] & r = 
        let Sheet(sheet1) = sheet;
        apply1(ma, sheet1)

    pub def point(x: a): SheetReader[r, a] = 
        SheetReader(_ -> Ok(x) as & r)


    pub def throwError(msg: String): SheetReader[r, a] =
        SheetReader(_ -> 
            Err(SheetIO/Internal/EvalError.emptyEvalError({source = msg, exnMessage = ""})) as & r
        )
    
    def throwEvalError(e: EvalError): SheetReader[r, a] =
        SheetReader(_ -> 
            Err(e) as & r
        )


    pub def swapError(ma: SheetReader[r, a], msg: String): SheetReader[r, a] =
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a)    => Ok(a)
                case Err(err) => Err({source = msg | err})
            }
        )


    pub def catchError(ma: SheetReader[r, a], catcher: EvalError -> SheetReader[r, a] & ef): SheetReader[r, a] & ef=
        SheetReader(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(e) => apply1(catcher(e) as & r, wb)
        }) as & ef


    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: SheetReader[r, a]): SheetReader[r, b] & ef = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => { let ans = fn(a) as & r; Ok(ans) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: SheetReader[r, a -> b & ef], ma: SheetReader[r, a]): SheetReader[r, b] & ef = 
        SheetReader(wb -> 
            match apply1(mf, wb) { 
                case Ok(f) => match apply1(ma, wb) { 
                    case Ok(a) => { let ans = f(a) as & r; Ok(ans) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(fn: a -> SheetReader[r, b] & ef, ma: SheetReader[r, a]): SheetReader[r, b] & ef = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => apply1(fn(a) as & r, wb)
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def orelse(ma: SheetReader[r, a], mb: SheetReader[r, a]): SheetReader[r, a] =
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a)  => Ok(a)
                case Err(_) => apply1(mb, wb)
            }
        )


    pub def testM(ma: SheetReader[r, a]): SheetReader[r, Bool] =
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(_)  => Ok(true)
                case Err(_) => Ok(false)
            }
        )

    ///
    /// Expose the result of operation `ma`.
    ///
    /// By definition this never fails.
    ///
    pub def result(ma: SheetReader[r, a]): SheetReader[r, Result[a, String]] = 
        SheetReader(wb -> 
            let ans = apply1(ma, wb) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
            Ok(ans)
        )

    pub def guard(x: Bool): SheetReader[r, Unit] = 
        match x { 
            case true => point()
            case false => throwError("guard")
        }






    def liftPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Sheet -> Result[a, PrimError] & ef): SheetReader[r, a] & ef = 
        SheetReader(s -> 
            let r = () as Region[r];
            {fn(r, s) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.emptyEvalError)
        )  as & ef


    pub def liftAction(fn: Unit -> a & ef): SheetReader[r, a] & ef = 
        SheetReader(_ -> { let a = fn() as & r; Ok(a)}) as & ef


    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): SheetReader[r, a] & ef = 
        SheetReader(_ -> 
            {fn() as & r} |> Result.mapErr(s -> SheetIO/Internal/EvalError.emptyEvalError({ source = s, exnMessage = ""}))
        ) as & ef

    pub def liftOption(x: Option[a], errMsg: String): SheetReader[r, a] = match x { 
        case Some(x1) => point(x1)
        case None     => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): SheetReader[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): SheetReader[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetReader[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    // SheetReader
 
    /// TODO - the SheetReader monad is dealing with both Sheet and Workbook functions.
    /// This was useful at first but is probably not the right abstraction...

    pub def getSheetName(): SheetReader[r, String] \ Read(r) =
        liftPrim(SheetIO/Internal/Sheet.getSheetName)

    pub def getFirstRowNum(): SheetReader[r, Int32] \ Read(r) =
        liftPrim(SheetIO/Internal/Sheet.getFirstRowNum)

    /// This may be higher than expected if rows at the end of the sheet have previously had content 
    /// which has been beleted.
    pub def getLastRowNum(): SheetReader[r, Int32] \ Read(r) =
        liftPrim(SheetIO/Internal/Sheet.getLastRowNum)


    pub def getPhysicalNumberOfRows(): SheetReader[r, Int32] \ Read(r) =
        liftPrim(SheetIO/Internal/Sheet.getPhysicalNumberOfRows)


    pub def getRowAt(ix: Int32, eval: RowEvaluator[r, a]): SheetReader[r, a] \ Read(r) =
        let RowEvaluator(eval1) = eval;
        let* primrow = liftPrim(r -> SheetIO/Internal/Sheet.getRow(r, ix));
        let* fmt     = liftPrim((r, _) -> SheetIO/DataFormatter.new(r) |> Ok);
        SheetReader(_ -> eval1(fmt, primrow))


    ///
    /// Working with Iterator<Row> - wrapped here as a Stream - seems safer from errors 
    /// than trying to use `getLastRowNum` to find the end.
    ///
    def getRowIterator(): SheetReader[r, Iterator[Row[r], r]] \ Read(r) =
        import org.apache.poi.ss.usermodel.Sheet.rowIterator(): ##java.util.Iterator & r;
        let* iter = liftPrim((_, s) -> rowIterator(s) |> Ok);
        Interop/Collections/RawIterator.new(() as Region[r], iter) |> Interop/Collections/RawIterator.iterator |> point



    pub def getHeaders(headerIndex: Int32): SheetReader[r, List[String]] \ Read(r) = region r1 {
        getRowAt(headerIndex, SheetIO/RowEvaluator.getRowContents(r1)) |> map(Array.toList)
    }



    /// This might be expressive enough to write all `RowEvaluator` traversals with...
    def foldRowsTraversal(range: {start :: Int32, end :: Int32}, 
                            skipErrors: Bool,
                            f: (b, a, Int32) -> b & ef, 
                            s: b, 
                            eval: RowEvaluator[r, a]): SheetReader[r, b] \ { Read(r), ef } =
        let RowEvaluator(eval1) = eval;
        let* fmt            = liftPrim((r, _) -> SheetIO/DataFormatter.new(r) |> Ok);
        let* rowIter        = getRowIterator();
        def loop(s1) = match Iterator.next(rowIter) {
            case None => point(s1)
            case Some(row1) => {
                let primRow = SheetIO/Common.getRow(row1);
                let* rownum = liftPrim((r, _) -> SheetIO/Internal/Row.getRowNum(r, primRow) |> Ok);
                if (rownum >= range.start and rownum < range.end) 
                    match eval1(fmt, primRow) {
                        case Err(e1) => if (not skipErrors) throwEvalError(e1) else loop(s1)
                        case Ok(ans) => loop(f(s1, ans, rownum))
                    }
                else 
                    loop(s1)
            }
        };
        loop(s)

    /// This might be expressive enough to write all `RowEvaluator` traversals with...
    def foldNamedRowsTraversal(range: {start :: Int32, end :: Int32}, 
                                skipErrors: Bool,
                                headers: List[String],
                                f: (b, a, Int32) -> b & ef, 
                                s: b, 
                                eval: NamedRowEvaluator[r, a]): SheetReader[r, b] \ { Read(r), ef } =
        let NamedRowEvaluator(eval1) = eval;
        let ixs             = SheetIO/NamedRowEvaluator.makeHeaders(headers);        
        let* fmt            = liftPrim((r, _) -> SheetIO/DataFormatter.new(r) |> Ok);
        let* rowIter        = getRowIterator();
        def loop(s1) = match Iterator.next(rowIter) {
            case None => point(s1)
            case Some(row1) => {
                let primRow = SheetIO/Common.getRow(row1);
                let* rownum = liftPrim((r, _) -> SheetIO/Internal/Row.getRowNum(r, primRow) |> Ok);
                if (rownum >= range.start and rownum < range.end) 
                    match eval1(ixs, fmt, primRow) {
                        case Err(e1) => if (not skipErrors) throwEvalError(e1) else loop(s1)
                        case Ok(ans) => loop(f(s1, ans, rownum))
                    }
                else 
                    loop(s1)
            }
        };
        loop(s)



    pub def allRows(eval: RowEvaluator[r, a]): SheetReader[r, List[a]] \ { Read(r) } =
        let* end = getPhysicalNumberOfRows();
        foldRowsTraversal({start = 0, end = end}, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    /// rows where `eval` fails are discounted...
    pub def someRows(eval: RowEvaluator[r, a]): SheetReader[r, List[a]] \ { Read(r) } =
        let* end = getPhysicalNumberOfRows();
        foldRowsTraversal({start = 0, end = end}, true, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    /// WARNING - start will be "miscalculated" if there are blank rows between 0 and start.
    /// We need a more robust implementation than this...

    pub def allRowsFrom(start: Int32, eval: RowEvaluator[r, a]): SheetReader[r, List[a]] \ { Read(r) } =
        let* end = getPhysicalNumberOfRows();
        foldRowsTraversal({start = start, end = end}, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    pub def rowsInRange(range: {start :: Int32, end :: Int32}, eval: RowEvaluator[r, a]): SheetReader[r, List[a]] \ { Read(r) } =
        foldRowsTraversal(range, false, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))


    /// Headers at 0, start at 1
    pub def allNamedRows(eval: NamedRowEvaluator[r, a]): SheetReader[r, List[a]] \ { Read(r) } =
        let* headers = getHeaders(0);
        let* end = getPhysicalNumberOfRows();
        foldNamedRowsTraversal({start = 1, end = end}, false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))

    pub def allNamedRowsFrom(pos: {headerIndex :: Int32, start :: Int32}, eval: NamedRowEvaluator[r, a]): SheetReader[r, List[a]] \ { Read(r) } =
        let* headers = getHeaders(pos.headerIndex);
        let* end = getPhysicalNumberOfRows();
        foldNamedRowsTraversal({start = pos.start, end = end}, false, headers, (k, a, _) -> ks -> k(a :: ks), ks -> ks, eval) |> map(k -> k(Nil))



/*


    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someRows(eval: RowEvaluator[r, Option[a]], sheet: Sheet): SheetReader[r, List[a]] = 
        let postprocess = xs -> List.foldRightWithCont((x, fac) -> match x {
                case None => fac()
                case Some(x1) => x1 :: fac()
            }, Nil, xs);
        map(postprocess, allRows(eval, sheet))


    

    pub def rowEval(fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                    eval: RowEvaluator[r, a], 
                    x: Row): SheetReader[r, a] =
        let RowEvaluator(f) = eval;
        let Row(x1) = x;
        liftActionResult(_ -> f(fmt, x1) |> Result.mapErr(SheetIO/EvalError.toString))



    pub def namedRowsInRange(ixs: Map[String, Int32], b: Int32, e: Int32, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, List[a]] = 
        let NamedRowEvaluator(f) = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= b and ix < e) Some(row) else None
        };
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* xs             = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.toList!);
        point(xs)

    pub def nameRowEval(indices: Map[String, Int32], 
                        fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                        eval: NamedRowEvaluator[r, a], x: Row): SheetReader[r, a] =
        let NamedRowEvaluator(f) = eval;
        let Row(x1) = x;
        liftActionResult(_ -> f(indices, fmt, x1) |> Result.mapErr(SheetIO/EvalError.toString))





    /// Reader headers from row at `headerIndex` then read data rows from `start`.
    /// Note - might be wise to change back to flix-sandbox Cursor while stdlib Interator
    /// is in development.
    /// `headerIndex` and `start` are zero indexed (as per POI rather than Excel).
    ///
    pub def namedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, List[a]] =   
        let NamedRowEvaluator(f)    = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= r.start) Some(row) else None
        };
        let* fmt        = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* row0       = getRowAt(sheet, r.headerIndex);
        let* ixs        = region r1 {
            let* hdrs = rowEval(fmt, SheetIO/RowEvaluator.getRowContents(r1), row0);
            point(SheetIO/NamedRowEvaluator.makeHeaders(hdrs))
        };
        let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* xs         = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.toList!);
        point(xs)


    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allNamedRows(eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, List[a]] = 
        namedRowsFrom({headerIndex = 0, start = 1}, eval, sheet)

    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someNamedRows(eval: NamedRowEvaluator[r, Option[a]], sheet: Sheet): SheetReader[r, List[a]] = 
        let postprocess = xs -> List.foldRightWithCont((x, fac) -> match x {
                case None => fac()
                case Some(x1) => x1 :: fac()
            }, Nil, xs);
        map(postprocess, allNamedRows(eval, sheet))


    pub def foldNamedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, 
                                combine: (b, a) -> b & ef, ac: b, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, b] =
        let NamedRowEvaluator(f) = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= r.start) Some(row) else None
        };
        let* fmt        = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* row0       = getRowAt(sheet, r.headerIndex);
        let* ixs        = region r1 { 
            let* hdrs = rowEval(fmt, SheetIO/RowEvaluator.getRowContents(r1), row0);
            point(SheetIO/NamedRowEvaluator.makeHeaders(hdrs))
        };
        let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* ans        = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.fold!(combine, ac));
        point(ans)


    pub def foldAllNamedRows(combine: (b, a) -> b & ef, ac: b, eval: NamedRowEvaluator[r, a], sheet: Sheet): SheetReader[r, b] =
        foldNamedRowsFrom({headerIndex = 0, start = 1}, combine, ac, eval, sheet)


    ///
    /// Working with Iterator<Row> - wrapped here as a Stream - seems safer from errors 
    /// than trying to use `getLastRowNum` to find the end.
    ///
    def getRowCursor(x: Sheet): Cursor[String, ##org.apache.poi.ss.usermodel.Row] & Impure =
        import new flixspt.sheetio.POIRowIterator(##org.apache.poi.ss.usermodel.Sheet): ##flixspt.sheetio.POIRowIterator & Impure as newRowCursor;
        import flixspt.sheetio.POIRowIterator.hasNext(): Bool & Impure as ffiHasNext;
        import flixspt.sheetio.POIRowIterator.next(): ##org.apache.poi.ss.usermodel.Row & Impure as ffiNext;
        try {
            let Sheet(x1)       = x;
            let iter = newRowCursor(x1);
            let cursorAlg =  {
                  hasNext   = () -> ffiHasNext(iter)
                , next      = FlixSandbox/Deprecated/Cursor.guardThunk(_ -> ffiNext(iter), "next failed")
                , finalize  = () -> () as & Impure
            };
            FlixSandbox/Deprecated/Cursor.makeCursor(cursorAlg)
        } catch {
            case _: ##java.lang.Exception => FlixSandbox/Deprecated/Cursor.failure("newPOIRowStream")
        }

    pub def mapWithIndexM(mf: (a, Int32) -> SheetReader[r, b], xs: List[a]): SheetReader[r, List[b]] =
        SheetReader(wb ->
            mapWithIndexMHelper(wb, mf, 0, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            mf: (a, Int32) -> SheetReader[r, b], 
                            ix : Int32, 
                            xs: List[a], 
                            fk: String -> Result[List[b], String], 
                            sk: List[b] -> Result[List[b], String]): Result[List[b], String] & Impure =
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

    pub def mapWithIndexX(mf: (a, Int32) -> SheetReader[r, b], xs: List[a]): SheetReader[r, Unit] =
        SheetReader(wb ->
            mapWithIndexXHelper(wb, mf, 0, xs, e -> Err(e), _ -> Ok())
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            mf: (a, Int32) -> SheetReader[r, b],
                            ix: Int32, 
                            xs: List[a], 
                            fk: String -> Result[Unit, String], 
                            sk: Unit -> Result[Unit, String]): Result[Unit, String] & Impure =
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
    pub def traverse(f: a -> SheetReader[r, b], xs: List[a]): SheetReader[r, List[b]] = mapM(f, xs)

    ///
    /// `traverseX` is a synonym for `mapX`.
    ///
    pub def traverseX(f: a -> SheetReader[r, b], xs: List[a]): SheetReader[r, Unit] = mapX(f, xs)

    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    pub def forM(xs: List[a], f: a -> SheetReader[r, b]): SheetReader[r, List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    pub def forX(xs: List[a], f: a -> SheetReader[r, b]): SheetReader[r, Unit] = mapX(f, xs)


    pub def foldM(mf: (b, a) -> SheetReader[r, b], start: b, xs: List[a]): SheetReader[r, b] =
        SheetReader(wb ->
            foldMHelper(wb, mf, start, xs, e -> Err(e), ac -> Ok(ac))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    mf: (b, a) -> SheetReader[r, b], 
                    acc: b, 
                    xs: List[a], 
                    fk: String -> Result[b, String], 
                    sk: b -> Result[b, String]): Result[b, String] & Impure =
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

    pub def mapAccumM(mf: (ac, a) -> SheetReader[r, (ac, b)], start: ac, xs: List[a]): SheetReader[r, (ac, List[b])] =
        SheetReader(wb ->
            mapAccumMHelper(wb, mf, start, xs, e -> Err(e), (ac, ks) -> Ok((ac, ks)))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        mf: (ac, a) -> SheetReader[r, (ac, b)],
                        ac: ac,
                        xs: List[a],
                        fk: String -> Result[(ac, List[b]), String],
                        sk: (ac, List[b]) -> Result[(ac, List[b]), String]): Result[(ac, List[b]), String] & Impure =
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
    pub def mapAccumX(mf: (ac, a) -> SheetReader[r, ac], start: ac, xs: List[a]): SheetReader[r, ac] = foldM(mf, start, xs)


*/

}
