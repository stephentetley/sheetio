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

    pub enum SheetReader[a](##org.apache.poi.ss.usermodel.Workbook -> Result[a, String] & Impure)


    instance Functor[SheetReader] {
        pub def map(f: a -> b & ef, o: SheetReader[a]): SheetReader[b] & ef = 
            SheetIO/SheetReader.map(f, o)
    }

    instance Applicative[SheetReader] {
        pub def point(x: a): SheetReader[a] = 
            SheetIO/SheetReader.point(x)

        pub def ap(f: SheetReader[a -> b & ef], x: SheetReader[a]) : SheetReader[b] & ef = 
            SheetIO/SheetReader.ap(f, x)
    }

    instance Monad[SheetReader] {
        pub def flatMap(f: a -> SheetReader[b] & ef, o: SheetReader[a]): SheetReader[b] & ef =
            SheetIO/SheetReader.flatMap(f, o)
    }

}

namespace SheetIO/SheetReader {

    use ToString.toString;

    use FlixSandbox/Deprecated.Cursor;

    use SheetIO.SheetReader;
    use SheetIO.SheetReader.{SheetReader};

    use SheetIO.Row;
    use SheetIO.Row.{Row};
    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.Sheet;
    use SheetIO.Sheet.{Sheet};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};


    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: SheetReader[a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(ma: SheetReader[a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{openXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let path1 = toString(path);
        match openXSSFWorkbook(path1) { 
            case Err(e) => Err("runXssf('${path1}'): ${e}")
            case Ok(wb) => {
                let ans = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook);
                let* _ = closeXSSFWorkbook(wb);
                ans         
            }
        }

    /// Excel 97 binary (.xls) format.
    pub def runHssf(ma: SheetReader[a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{openHSSFWorkbook, closeHSSFWorkbook};
        use Result.{flatMap};
        let path1 = toString(path);
        match openHSSFWorkbook(path1) { 
            case Err(_) => Err("runHssf('${path1}')")
            case Ok(wb) => {
                let ans = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook);
                let* _ = closeHSSFWorkbook(wb);
                ans         
            }
        }


    def apply1(ma: SheetReader[a], wb: ##org.apache.poi.ss.usermodel.Workbook): Result[a, String] & Impure = 
        let SheetReader(x1) = ma;
        x1(wb)

    pub def point(x: a): SheetReader[a] = 
        SheetReader(_ -> Ok(x) as & Impure)

    pub def throwError(msg: String): SheetReader[a] =
        SheetReader(_ -> Err(msg) as & Impure)


    pub def swapError(ma: SheetReader[a], msg: String): SheetReader[a] =
        SheetReader(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(_) => Err(msg)
        })


    pub def catchError(ma: SheetReader[a], handle: String -> SheetReader[a]): SheetReader[a] =
        SheetReader(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(e) => apply1(handle(e), wb)
        })

    pub def mapError(f: String -> String, ma: SheetReader[a]): SheetReader[a] =
        SheetReader(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(e) => Err(f(e))
        })

    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: SheetReader[a]): SheetReader[b] & ef = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => { let ans = fn(a) as & Impure; Ok(ans) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: SheetReader[a -> b & ef], ma: SheetReader[a]): SheetReader[b] & ef = 
        SheetReader(wb -> 
            match apply1(mf, wb) { 
                case Ok(f) => match apply1(ma, wb) { 
                    case Ok(a) => { let ans = f(a) as & Impure; Ok(ans) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def bind(ma: SheetReader[a], fn: a -> SheetReader[b] & ef): SheetReader[b] & ef = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => apply1(fn(a), wb)
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(fn: a -> SheetReader[b] & ef, ma: SheetReader[a]): SheetReader[b] & ef = 
        bind(ma, fn)


    pub def alt(ma: SheetReader[a], mb: SheetReader[a]): SheetReader[a] =
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => Ok(a)
                case Err(_) => apply1(mb, wb)
            }
        )

    pub def optional(ma: SheetReader[a]): SheetReader[Option[a]] = 
        alt(map(Some, ma), point(None))

    pub def tryWithDefault(ma: SheetReader[a], d: a): SheetReader[a] = 
        alt(ma, point(d))

    pub def testM(ma: SheetReader[a]): SheetReader[Bool] =
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(_) => Ok(true)
                case Err(_) => Ok(false)
            }
        )

    ///
    /// Expose the result of operation `ma`.
    ///
    /// By definition this never fails.
    ///
    pub def result(ma: SheetReader[a]): SheetReader[Result[a, String]] = 
        SheetReader(wb -> 
            let ans = apply1(ma, wb);
            Ok(ans)
        )

    pub def guard(x: Bool): SheetReader[Unit] = 
        match x { 
            case true => point()
            case false => throwError("guard")
        }

    pub def whenTrue(x: Bool, ma: SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case true => ma
            case false => point()
        }

    pub def whenFalse(x: Bool, ma: SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case true => point()
            case false => ma
        }

    pub def whenTrueM(mx: SheetReader[Bool], ma: SheetReader[Unit]): SheetReader[Unit] = 
        bind(mx, x -> whenTrue(x, ma))

    pub def whenFalseM(mx: SheetReader[Bool], ma: SheetReader[Unit]): SheetReader[Unit] = 
        bind(mx, x -> whenFalse(x, ma))


    pub def whenSome(x: Option[a], f: a -> SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case Some(a) => f(a)
            case None => point()
        }

    pub def whenNone(x: Option[a], ma: SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case Some(_) => point()
            case None => ma
        }

    pub def whenSomeM(mx: SheetReader[Option[a]], f: a -> SheetReader[Unit]): SheetReader[Unit] = 
        bind(mx, x -> whenSome(x, f))
            

    pub def whenNoneM(mx: SheetReader[Option[a]], ma: SheetReader[Unit]): SheetReader[Unit] = 
        bind(mx, x -> whenNone(x, ma))


    pub def ifSomeM(mx: SheetReader[Option[a]], mthen: a -> SheetReader[b], melse: SheetReader[b]): SheetReader[b] = 
        bind(mx, x -> match x { 
            case Some(a) => mthen(a)
            case None => melse
        })
            

    pub def ifNoneM(mx: SheetReader[Option[a]], mthen: SheetReader[b], melse: SheetReader[b]): SheetReader[b] = 
        bind(mx, x -> match x { 
            case Some(_) => mthen
            case None => melse
        })

    def liftWorkbookPrim(fn: ##org.apache.poi.ss.usermodel.Workbook -> Result[a, String] & Impure): SheetReader[a] = 
        SheetReader(wb -> fn(wb))

    pub def liftAction(fn: Unit -> a & Impure): SheetReader[a] = 
        SheetReader(_ -> { let a = fn(); Ok(a) })

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): SheetReader[a] = 
        SheetReader(_ -> fn())

    pub def liftOption(x: Option[a], errMsg: String): SheetReader[a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): SheetReader[a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): SheetReader[a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetReader[a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    pub def mapM(mf: a -> SheetReader[b], xs: List[a]): SheetReader[List[b]] =
        SheetReader(wb ->
            mapMHelper(wb, mf, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   mf: a -> SheetReader[b], 
                   xs: List[a], 
                   fk: String -> Result[List[b], String], 
                   sk: List[b] -> Result[List[b], String]): Result[List[b], String] & Impure =
        match xs {
            case Nil => sk(Nil)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb) {
                    case Err(e) => fk(e)
                    case Ok(x1) => mapMHelper(wb, mf, rs, fk, ks -> sk(x1 :: ks))
                }
            }
        }

    pub def mapX(mf: a -> SheetReader[b], xs: List[a]): SheetReader[Unit] =
        SheetReader(wb ->
            mapXHelper(wb, mf, xs, e -> Err(e), _ -> Ok())
        )

    def mapXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   mf: a -> SheetReader[b], 
                   xs: List[a], 
                   fk: String -> Result[Unit, String], 
                   sk: Unit -> Result[Unit, String]): Result[Unit, String] & Impure =
        match xs {
            case Nil => sk()
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb) {
                    case Err(e) => fk(e)
                    case Ok(_) => mapXHelper(wb, mf, rs, fk, sk)
                }
            }
        }

    pub def mapWithIndexM(mf: (a, Int32) -> SheetReader[b], xs: List[a]): SheetReader[List[b]] =
        SheetReader(wb ->
            mapWithIndexMHelper(wb, mf, 0, xs, e -> Err(e), ys -> Ok(ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            mf: (a, Int32) -> SheetReader[b], 
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

    pub def mapWithIndexX(mf: (a, Int32) -> SheetReader[b], xs: List[a]): SheetReader[Unit] =
        SheetReader(wb ->
            mapWithIndexXHelper(wb, mf, 0, xs, e -> Err(e), _ -> Ok())
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            mf: (a, Int32) -> SheetReader[b],
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


    pub def sequence(xs: List[SheetReader[a]]): SheetReader[List[a]] =
        SheetReader(wb ->
            sequenceHelper(wb, xs, e -> Err(e), ks -> Ok(ks))
        )

    ///
    /// Helper function for `sequence`.
    ///
    def sequenceHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                       xs: List[SheetReader[a]], 
                       fk: String -> Result[List[a], String], 
                       sk: List[a] -> Result[List[a], String]): Result[List[a], String] & Impure =
        match xs {
            case Nil => sk(Nil)
            case mx :: rs => match apply1(mx, wb) {
                case Err(e) => fk(e)
                case Ok(x) => sequenceHelper(wb, rs, fk, ks -> sk(x :: ks))
            }
        }

    ///
    /// Returns `(v1 :: v2 :: ... v :: vn)` if each of `f(xs_i)` is successful.
    /// Otherwise fails with the first error encountered.
    ///
    /// `traverse` is a synonym for `mapM`.
    ///
    pub def traverse(f: a -> SheetReader[b], xs: List[a]): SheetReader[List[b]] = mapM(f, xs)

    ///
    /// `traverseX` is a synonym for `mapX`.
    ///
    pub def traverseX(f: a -> SheetReader[b], xs: List[a]): SheetReader[Unit] = mapX(f, xs)

    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    pub def forM(xs: List[a], f: a -> SheetReader[b]): SheetReader[List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    pub def forX(xs: List[a], f: a -> SheetReader[b]): SheetReader[Unit] = mapX(f, xs)


    pub def foldM(mf: (b, a) -> SheetReader[b], start: b, xs: List[a]): SheetReader[b] =
        SheetReader(wb ->
            foldMHelper(wb, mf, start, xs, e -> Err(e), ac -> Ok(ac))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    mf: (b, a) -> SheetReader[b], 
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

    pub def mapAccumM(mf: (ac, a) -> SheetReader[(ac, b)], start: ac, xs: List[a]): SheetReader[(ac, List[b])] =
        SheetReader(wb ->
            mapAccumMHelper(wb, mf, start, xs, e -> Err(e), (ac, ks) -> Ok((ac, ks)))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        mf: (ac, a) -> SheetReader[(ac, b)],
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
    pub def mapAccumX(mf: (ac, a) -> SheetReader[ac], start: ac, xs: List[a]): SheetReader[ac] = foldM(mf, start, xs)



    // SheetReader

    pub def getNumberOfSheets(): SheetReader[Int32] =
        use SheetIO/Internal/Workbook.{getNumberOfSheets};
        liftWorkbookPrim(getNumberOfSheets)

    pub def getSheetNames(): SheetReader[List[String]] =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0,n);
        mapM(getSheetNameWb, ixs)      


    def getSheetNameWb(ix: Int32): SheetReader[String] =
        use SheetIO/Internal/Workbook.{getSheetName};
        liftWorkbookPrim(x1 -> getSheetName(x1, ix))


    pub def getSheetNamed(name: String): SheetReader[Sheet] =
        use SheetIO/Internal/Workbook.{getSheet};
        liftWorkbookPrim(x1 -> getSheet(x1, name) |> Result.map(Sheet)) 

    pub def getSheetAt(ix: Int32): SheetReader[Sheet] =
        use SheetIO/Internal/Workbook.{getSheetAt};
        liftWorkbookPrim(x1 -> getSheetAt(x1, ix) |> Result.map(Sheet))

    pub def getSheetName(x: Sheet): SheetReader[String] =
        use SheetIO/Internal/Sheet.{getSheetName};
        let Sheet(x1) = x;
        liftActionResult(_ -> getSheetName(x1))

    pub def getFirstRowNum(x: Sheet): SheetReader[Int32] =
        use SheetIO/Internal/Sheet.{getFirstRowNum};
        let Sheet(x1) = x;
        liftActionResult(_ -> getFirstRowNum(x1))

    /// This may be higher than expected if rows at the end of the sheet have previously had content 
    /// which has been beleted.
    pub def getLastRowNum(x: Sheet): SheetReader[Int32] =
        use SheetIO/Internal/Sheet.{getLastRowNum};
        let Sheet(x1) = x;
        liftActionResult(_ -> getLastRowNum(x1))


    pub def getPhysicalNumberOfRows(x: Sheet): SheetReader[Int32] =
        use SheetIO/Internal/Sheet.{getPhysicalNumberOfRows};
        let Sheet(x1) = x;
        liftActionResult(_ -> getPhysicalNumberOfRows(x1))


    pub def getRowAt(x: Sheet, ix: Int32): SheetReader[Row] =
        use SheetIO/Internal/Sheet.{getRow};
        let Sheet(x1) = x;
        liftActionResult(_ -> getRow(x1, ix) |> Result.map(Row))

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

    /// Use rowiterator...
    pub def allRows(eval: RowEvaluator[a], sheet: Sheet): SheetReader[List[a]] =         
        let RowEvaluator(f) = eval;
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet));
        let* xs             = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.toList!);
        point(xs)


    /// WARNING - start will be "miscalculated" if there are blank rows between 0 and start.
    /// We need a more robust implementation than this...
    pub def allRowsFrom(start: Int32, eval: RowEvaluator[a], sheet: Sheet): SheetReader[List[a]] =         
        let RowEvaluator(f) = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= start) Some(row) else None
        };
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* xs             = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.toList!);
        point(xs)


    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someRows(eval: RowEvaluator[Option[a]], sheet: Sheet): SheetReader[List[a]] = 
        let postprocess = xs -> List.foldRightWithCont((x, fac) -> match x {
                case None => fac()
                case Some(x1) => x1 :: fac()
            }, Nil, xs);
        map(postprocess, allRows(eval, sheet))

    pub def foldAllRows(combine: (b, a) -> b & ef, ac: b, eval: RowEvaluator[a], sheet: Sheet): SheetReader[b] =
        let RowEvaluator(f) = eval;
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet));
        let* ans            = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.fold!(combine, ac));
        point(ans)



    pub def rowsInRange(b: Int32, e: Int32, eval: RowEvaluator[a], sheet: Sheet): SheetReader[List[a]] = 
        let RowEvaluator(f) = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= b and ix < e) Some(row) else None
        };
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* xs             = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.toList!);
        point(xs)

    pub def rowEval(fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                    eval: RowEvaluator[a], 
                    x: Row): SheetReader[a] =
        let RowEvaluator(f) = eval;
        let Row(x1) = x;
        liftActionResult(_ -> f(fmt, x1) |> Result.mapErr(SheetIO/EvalError.toString))



    pub def namedRowsInRange(ixs: Map[String, Int32], b: Int32, e: Int32, eval: NamedRowEvaluator[a], sheet: Sheet): SheetReader[List[a]] = 
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
                        eval: NamedRowEvaluator[a], x: Row): SheetReader[a] =
        let NamedRowEvaluator(f) = eval;
        let Row(x1) = x;
        liftActionResult(_ -> f(indices, fmt, x1) |> Result.mapErr(SheetIO/EvalError.toString))





    /// Reader headers from row at `headerIndex` then read data rows from `start`.
    /// Note - might be wise to change back to flix-sandbox Cursor while stdlib Interator
    /// is in development.
    /// `headerIndex` and `start` are zero indexed (as per POI rather than Excel).
    ///
    pub def namedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, eval: NamedRowEvaluator[a], sheet: Sheet): SheetReader[List[a]] =   
        let NamedRowEvaluator(f)    = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= r.start) Some(row) else None
        };
        let* fmt        = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* row0       = getRowAt(sheet, r.headerIndex);
        let* hdrs       = rowEval(fmt, SheetIO/RowEvaluator.getRowContents(), row0);
        let* ixs        = liftAction(_ -> SheetIO/NamedRowEvaluator.makeHeaders(hdrs));
        let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* xs         = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.toList!);
        point(xs)


    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allNamedRows(eval: NamedRowEvaluator[a], sheet: Sheet): SheetReader[List[a]] = 
        namedRowsFrom({headerIndex = 0, start = 1}, eval, sheet)

    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someNamedRows(eval: NamedRowEvaluator[Option[a]], sheet: Sheet): SheetReader[List[a]] = 
        let postprocess = xs -> List.foldRightWithCont((x, fac) -> match x {
                case None => fac()
                case Some(x1) => x1 :: fac()
            }, Nil, xs);
        map(postprocess, allNamedRows(eval, sheet))


    pub def foldNamedRowsFrom(r: {headerIndex :: Int32, start :: Int32}, 
                                combine: (b, a) -> b & ef, ac: b, eval: NamedRowEvaluator[a], sheet: Sheet): SheetReader[b] =
        let NamedRowEvaluator(f) = eval;
        let checker         = row -> {
            let ix = SheetIO/Internal/Row.getRowNum(row);
            if (ix >= r.start) Some(row) else None
        };
        let* fmt        = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* row0       = getRowAt(sheet, r.headerIndex);
        let* hdrs       = rowEval(fmt, SheetIO/RowEvaluator.getRowContents(), row0);
        let* ixs        = liftAction(_ -> SheetIO/NamedRowEvaluator.makeHeaders(hdrs));
        let* strm       = liftAction(_ -> getRowCursor(sheet) |> FlixSandbox/Deprecated/Cursor.filterMap(checker));
        let* ans        = liftActionResult(_ -> FlixSandbox/Deprecated/Cursor.mapResult(f(ixs, fmt) >> Result.mapErr(SheetIO/EvalError.toString), strm) |> FlixSandbox/Deprecated/Cursor.fold!(combine, ac));
        point(ans)

    pub def foldAllNamedRows(combine: (b, a) -> b & ef, ac: b, eval: NamedRowEvaluator[a], sheet: Sheet): SheetReader[b] =
        foldNamedRowsFrom({headerIndex = 0, start = 1}, combine, ac, eval, sheet)

}
