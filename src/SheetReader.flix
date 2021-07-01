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

    pub opaque type SheetReader[a] = ##org.apache.poi.ss.usermodel.Workbook ~> Result[a, String]

}

namespace SheetIO/SheetReader {

    use ToString.toString;
    
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
    pub def runXlsx(ma: SheetReader[a], path: System.FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(ma: SheetReader[a], path: System.FilePath): Result[a, String] & Impure = 
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
    pub def runHssf(ma: SheetReader[a], path: System.FilePath): Result[a, String] & Impure = 
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

    pub def return(x: a): SheetReader[a] = 
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
    pub def map(fn: a -> b & ef, ma: SheetReader[a]): SheetReader[b] = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => { let ans = fn(a) as & Impure; Ok(ans) }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: SheetReader[a -> b & ef], ma: SheetReader[a]): SheetReader[b] = 
        SheetReader(wb -> 
            match apply1(mf, wb) { 
                case Ok(f) => match apply1(ma, wb) { 
                    case Ok(a) => { let ans = f(a) as & Impure; Ok(ans) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqLeft(ma: SheetReader[a], mb: SheetReader[b]): SheetReader[a] = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => match apply1(mb, wb) {
                    case Ok(_) => Ok(a)
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqRight(ma: SheetReader[a], mb: SheetReader[b]): SheetReader[b] = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(_) => apply1(mb, wb)
                case Err(e) => Err(e)
            }
        )


    pub def bind(ma: SheetReader[a], fn: a -> SheetReader[b]): SheetReader[b] = 
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => apply1(fn(a), wb)
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(fn: a -> SheetReader[b], ma: SheetReader[a]): SheetReader[b] = bind(ma, fn)

    pub def kleisliLeft(f1: a -> SheetReader[b], f2: b -> SheetReader[c], x: a): SheetReader[c] =
        SheetReader(wb -> 
            match apply1(f1(x), wb) { 
                case Ok(b) => apply1(f2(b), wb)
                case Err(e) => Err(e)
            }
        )

    pub def kleisliRight(f1: b -> SheetReader[c], f2: a -> SheetReader[b], x: a): SheetReader[c] =
        kleisliLeft(f2, f1, x)


    pub def alt(ma: SheetReader[a], mb: SheetReader[a]): SheetReader[a] =
        SheetReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => Ok(a)
                case Err(_) => apply1(mb, wb)
            }
        )

    pub def optional(ma: SheetReader[a]): SheetReader[Option[a]] = 
        alt(map(Some, ma), return(None))

    pub def tryWithDefault(ma: SheetReader[a], d: a): SheetReader[a] = 
        alt(ma, return(d))

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
            case true => return()
            case false => throwError("guard")
        }

    pub def whenTrue(x: Bool, ma: SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case true => ma
            case false => return()
        }

    pub def whenFalse(x: Bool, ma: SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case true => return()
            case false => ma
        }

    pub def whenTrueM(mx: SheetReader[Bool], ma: SheetReader[Unit]): SheetReader[Unit] = 
        bind(mx, x -> whenTrue(x, ma))

    pub def whenFalseM(mx: SheetReader[Bool], ma: SheetReader[Unit]): SheetReader[Unit] = 
        bind(mx, x -> whenFalse(x, ma))


    pub def whenSome(x: Option[a], f: a -> SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case Some(a) => f(a)
            case None => return()
        }

    pub def whenNone(x: Option[a], ma: SheetReader[Unit]): SheetReader[Unit] = 
        match x { 
            case Some(_) => return()
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

    def liftWorkbookPrim(fn: ##org.apache.poi.ss.usermodel.Workbook ~> Result[a, String]): SheetReader[a] = 
        SheetReader(wb -> fn(wb))

    pub def liftAction(fn: Unit ~> a): SheetReader[a] = 
        SheetReader(_ -> { let a = fn(); Ok(a) })

    pub def liftActionResult(fn: Unit ~> Result[a, String]): SheetReader[a] = 
        SheetReader(_ -> fn())

    pub def liftOption(x: Option[a], errMsg: String): SheetReader[a] = match x { 
        case Some(x1) => return(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): SheetReader[a] = match x { 
        case Ok(x1) => return(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetReader[a] = match x { 
        case Success(x1) => return(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
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

    /// Alias for `seqRight`
    pub def *>(p1: SheetReader[a], p2: SheetReader[b]): SheetReader[b] = seqRight(p1, p2)

    /// Alias for `seqLeft`
    pub def <*(p1: SheetReader[a], p2: SheetReader[b]): SheetReader[a] = seqLeft(p1, p2)

    /// Haskell's <|>
    pub def <!>(p1: SheetReader[a], p2: SheetReader[a]): SheetReader[a] =
        alt(p1, p2)

    // Haskell's <$>
    pub def <&>(f: a -> b, p2: SheetReader[a]): SheetReader[b] =
        map(f, p2)

    pub def <&(x: b, p2: SheetReader[a]): SheetReader[b] =
        p2 *> return(x)

    /// Alias for `bind`
    pub def >>=(ma: SheetReader[a], k: a -> SheetReader[b]): SheetReader[b] = bind(ma, k)

    /// Alias for `flatMap` (flipped `bind`)
    pub def =<<(k: a -> SheetReader[b], ma: SheetReader[a]): SheetReader[b] = flatMap(k, ma)


    // Haskell's ap <*>
    pub def <*>(mf: SheetReader[a -> b], p2: SheetReader[a]): SheetReader[b] =
        ap(mf, p2)

    /// Alias for `kleisliLeft`
    pub def >=>(f1: a -> SheetReader[b], f2: b -> SheetReader[c]): a -> SheetReader[c] = x -> kleisliLeft(f1, f2, x)

    /// Alias for `kleisliRight`
    pub def <=<(f1: b -> SheetReader[c], f2: a -> SheetReader[b]): a -> SheetReader[c] = x -> kleisliRight(f1, f2, x)


    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2]): SheetReader[u] =
        ap(map(f, o1), o2)


    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3]): SheetReader[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3], 
                  o4: SheetReader[t4]): SheetReader[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3], 
                  o4: SheetReader[t4], 
                  o5: SheetReader[t5]): SheetReader[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3],
                  o4: SheetReader[t4], 
                  o5: SheetReader[t5], 
                  o6: SheetReader[t6]): SheetReader[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3], 
                  o4: SheetReader[t4], 
                  o5: SheetReader[t5], 
                  o6: SheetReader[t6], 
                  o7: SheetReader[t7]): SheetReader[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3], 
                  o4: SheetReader[t4], 
                  o5: SheetReader[t5], 
                  o6: SheetReader[t6], 
                  o7: SheetReader[t7], 
                  o8: SheetReader[t8]): SheetReader[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef, 
                  o1: SheetReader[t1], 
                  o2: SheetReader[t2], 
                  o3: SheetReader[t3], 
                  o4: SheetReader[t4], 
                  o5: SheetReader[t5], 
                  o6: SheetReader[t6], 
                  o7: SheetReader[t7], 
                  o8: SheetReader[t8], 
                  o9: SheetReader[t9]): SheetReader[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef, 
                   o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4], 
                   o5: SheetReader[t5], 
                   o6: SheetReader[t6], 
                   o7: SheetReader[t7], 
                   o8: SheetReader[t8], 
                   o9: SheetReader[t9], 
                   o10: SheetReader[t10]): SheetReader[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)


    ///
    /// Returns a 2 element tuple with the results of applying the evaluators `o1` and `o2`.
    ///
    /// Returns an error if evaluating either `o1` or `o2` is an error.
    ///
    pub def tuple2(o1: SheetReader[t1], 
                   o2: SheetReader[t2]): SheetReader[(t1, t2)] =
        lift2((x,y) -> (x,y), o1, o2)

    ///
    /// Returns a 3 element tuple with the results of applying the evaluators `o1`, `o2` and `o3`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2` and `o3` is an error.
    ///
    pub def tuple3(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3]): SheetReader[(t1, t2, t3)] =
        lift3((x1, x2, x3) -> (x1, x2, x3), o1, o2, o3)

    ///
    /// Returns a 4 element tuple with the results of applying the evaluators `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, `o3` and `o4` is an error.
    ///
    pub def tuple4(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4]): SheetReader[(t1, t2, t3, t4)]  =
        lift4((x1, x2, x3, x4) -> (x1, x2, x3, x4), o1, o2, o3, o4)

    ///
    /// Returns a 5 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o5` is an error.
    ///
    pub def tuple5(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4], 
                   o5: SheetReader[t5]): SheetReader[(t1, t2, t3, t4, t5)]  =
        lift5((x1, x2, x3, x4, x5) -> (x1, x2, x3, x4, x5), o1, o2, o3, o4, o5)

    ///
    /// Returns a 6 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o6` is an error.
    ///
    pub def tuple6(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4], 
                   o5: SheetReader[t5], 
                   o6: SheetReader[t6]): SheetReader[(t1, t2, t3, t4, t5, t6)]  =
        lift6((x1, x2, x3, x4, x5, x6) -> (x1, x2, x3, x4, x5, x6), o1, o2, o3, o4, o5, o6)

    ///
    /// Returns a 7 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o7` is an error.
    ///
    pub def tuple7(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4], 
                   o5: SheetReader[t5], 
                   o6: SheetReader[t6], 
                   o7: SheetReader[t7]): SheetReader[(t1, t2, t3, t4, t5, t6, t7)]  =
        lift7((x1, x2, x3, x4, x5, x6, x7) -> (x1, x2, x3, x4, x5, x6, x7), 
                o1, o2, o3, o4, o5, o6, o7)

    ///
    /// Returns a 8 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o8` is an error.
    ///
    pub def tuple8(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4], 
                   o5: SheetReader[t5], 
                   o6: SheetReader[t6], 
                   o7: SheetReader[t7], 
                   o8: SheetReader[t8]): SheetReader[(t1, t2, t3, t4, t5, t6, t7, t8)]  =
        lift8((x1, x2, x3, x4, x5, x6, x7, x8) -> (x1, x2, x3, x4, x5, x6, x7, x8), 
                o1, o2, o3, o4, o5, o6, o7, o8)

    ///
    /// Returns a 9 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o9` is an error.
    ///
    pub def tuple9(o1: SheetReader[t1], 
                   o2: SheetReader[t2], 
                   o3: SheetReader[t3], 
                   o4: SheetReader[t4], 
                   o5: SheetReader[t5], 
                   o6: SheetReader[t6], 
                   o7: SheetReader[t7], 
                   o8: SheetReader[t8], 
                   o9: SheetReader[t9]): SheetReader[(t1, t2, t3, t4, t5, t6, t7, t8, t9)]  =
        lift9((x1, x2, x3, x4, x5, x6, x7, x8, x9) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9)

    ///
    /// Returns a 10 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o10` is an error.
    ///
    pub def tuple10(o1: SheetReader[t1], 
                    o2: SheetReader[t2], 
                    o3: SheetReader[t3], 
                    o4: SheetReader[t4], 
                    o5: SheetReader[t5], 
                    o6: SheetReader[t6], 
                    o7: SheetReader[t7], 
                    o8: SheetReader[t8], 
                    o9: SheetReader[t9], 
                    o10: SheetReader[t10]): SheetReader[(t1, t2, t3, t4, t5, t6, t7, t8, t9, t10)]  =
        lift10((x1, x2, x3, x4, x5, x6, x7, x8, x9, x10) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9, x10), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9, o10)

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
        import new flix.runtime.spt.sheetio.POIRowCursor(##org.apache.poi.ss.usermodel.Sheet) as newRowCursor;
        import flix.runtime.spt.sheetio.POIRowCursor.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.POIRowCursor.next() as ffiNext;
        import flix.runtime.spt.sheetio.POIRowCursor.close() as ffiClose;
        try {
            let Sheet(x1)       = x;
            let iter = newRowCursor(x1);
            let cursorAlg =  {
                  hasNext   = () -> ffiHasNext(iter)
                , next      = Cursor.guardThunk(_ -> ffiNext(iter), "next failed")
                , finalize  = () -> ffiClose(iter)
            };
            Cursor.makeCursor(cursorAlg)
        } catch {
            case _: ##java.lang.Exception => Cursor.failure("newPOIRowStream")
        }

    /// Use rowiterator...
    pub def allRows(sheet: Sheet, eval: RowEvaluator[a]): SheetReader[List[a]] =         
        let RowEvaluator(f) = eval;
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet));
        let* xs             = liftActionResult(_ -> Cursor.mapResult(f(fmt), strm) |> Cursor.toList!);
        return(xs)
    
    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someRows(sheet: Sheet, eval: RowEvaluator[a]): SheetReader[List[a]] = 
        let eval1 = SheetIO/RowEvaluator.alt(SheetIO/RowEvaluator.map(Some, eval), SheetIO/RowEvaluator.return(None));
        let postprocess = Option.sequence >> flip(Option.getWithDefault, Nil);
        map(postprocess, allRows(sheet, eval1))

    pub def foldAllRows(sheet: Sheet, combine: (b, a) -> b & ef, ac: b, eval: RowEvaluator[a]): SheetReader[b] =
        let RowEvaluator(f) = eval;
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet));
        let* ans            = liftActionResult(_ -> Cursor.mapResult(f(fmt), strm) |> Cursor.fold!(combine, ac));
        return(ans)



    pub def rowsInRange(sheet: Sheet, b: Int32, e: Int32, eval: RowEvaluator[a]): SheetReader[List[a]] = 
        let RowEvaluator(f) = eval;
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet));
        let* _              = liftAction(_ -> Cursor.drop!(b, strm));
        let* strm1          = liftAction(_ -> Cursor.take(e - b, strm));
        let* xs             = liftActionResult(_ -> Cursor.mapResult(f(fmt), strm1) |> Cursor.toList!);
        return(xs)

    pub def rowEval(x: Row, 
                    fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                    eval: RowEvaluator[a]): SheetReader[a] =
        let RowEvaluator(f) = eval;
        let Row(x1) = x;
        liftActionResult(_ -> f(fmt, x1))



    pub def namedRowsInRange(sheet: Sheet, ixs: Map[String, Int32], b: Int32, e: Int32, eval: NamedRowEvaluator[a]): SheetReader[List[a]] = 
        let NamedRowEvaluator(f) = eval;
        let* fmt            = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* strm           = liftAction(_ -> getRowCursor(sheet));
        let* _              = liftAction(_ -> Cursor.drop!(b, strm));
        let* strm1          = liftAction(_ -> Cursor.take(e - b, strm));
        let* xs             = liftActionResult(_ -> Cursor.mapResult(f(ixs, fmt), strm1) |> Cursor.toList!);
        return(xs)

    pub def nameRowEval(x: Row, indices: Map[String, Int32], 
                        fmt: ##org.apache.poi.ss.usermodel.DataFormatter, 
                        eval: NamedRowEvaluator[a]): SheetReader[a] =
        let NamedRowEvaluator(f) = eval;
        let Row(x1) = x;
        liftActionResult(_ -> f(indices, fmt, x1))





    /// Reader headers from row at `headerIndex` then read data rows from `start`.
    pub def namedRowsFrom(sheet: Sheet, headerIndex: Int32, start: Int32, eval: NamedRowEvaluator[a]): SheetReader[List[a]] =   
        let NamedRowEvaluator(f)    = eval;
        let* fmt        = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* r0         = getRowAt(sheet, headerIndex);
        let* hdrs       = rowEval(r0, fmt, SheetIO/RowEvaluator.getHeaders());
        let* ixs        = liftAction(_ -> SheetIO/NamedRowEvaluator.makeHeaders(hdrs));
        let* strm       = liftAction(_ -> getRowCursor(sheet));
        let* _          = liftAction(_ -> Cursor.drop!(start, strm));
        let* xs         = liftActionResult(_ -> Cursor.mapResult(f(ixs, fmt), strm) |> Cursor.toList!);
        return(xs)


    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allNamedRows(sheet: Sheet, eval: NamedRowEvaluator[a]): SheetReader[List[a]] = 
        namedRowsFrom(sheet, 0, 1, eval)

    /// Row 0 is read for the headers, some subsequent rows are *data*.
    pub def someNamedRows(sheet: Sheet, eval: NamedRowEvaluator[a]): SheetReader[List[a]] = 
        let eval1 = SheetIO/NamedRowEvaluator.alt(SheetIO/NamedRowEvaluator.map(Some, eval), SheetIO/NamedRowEvaluator.return(None));
        let postprocess = xs -> List.foldRight((x, ac) -> match x {
                case None => ac
                case Some(x1) => x1 :: ac
            }, Nil, xs);
        map(postprocess, allNamedRows(sheet, eval1))


    pub def foldNamedRowsFrom(sheet: Sheet, headerIndex: Int32, start: Int32, 
                                combine: (b, a) -> b & ef, ac: b, eval: NamedRowEvaluator[a]): SheetReader[b] =
        let NamedRowEvaluator(f) = eval;
        let* fmt        = liftAction(_ -> SheetIO/Internal/DataFormatter.new());
        let* r0         = getRowAt(sheet, headerIndex);
        let* hdrs       = rowEval(r0, fmt, SheetIO/RowEvaluator.getHeaders());
        let* ixs        = liftAction(_ -> SheetIO/NamedRowEvaluator.makeHeaders(hdrs));
        let* strm       = liftAction(_ -> getRowCursor(sheet));
        let* _          = liftAction(_ -> Cursor.drop!(start, strm));
        let* ans        = liftActionResult(_ -> Cursor.mapResult(f(ixs, fmt), strm) |> Cursor.fold!(combine, ac));
        return(ans)

    pub def foldAllNamedRows(sheet: Sheet, combine: (b, a) -> b & ef, ac: b, eval: NamedRowEvaluator[a]): SheetReader[b] =
        foldNamedRowsFrom(sheet, 0, 1, combine, ac, eval)

}
