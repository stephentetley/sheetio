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

    use SheetIO.CellStyle;
   
    pub opaque type SheetWriter[a] = 
        ##org.apache.poi.ss.usermodel.Workbook -> Option[CellStyle] -> Cursors ~> Result[(Cursors, a), String]

    type alias Cursors = Map[String, Int32]

}

namespace SheetIO/SheetWriter {

    use ToString.toString;

    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.CellStyleSetter;
    use SheetIO.CellStyleSetter.{CellStyleSetter};
    use SheetIO.DataFormat;
    use SheetIO.DataFormat.{DataFormat};
    use SheetIO.Font;
    use SheetIO.Font.{Font};
    use SheetIO.FontSetter;
    use SheetIO.FontSetter.{FontSetter};
    use SheetIO.Row;
    use SheetIO.Row.{Row};
    use SheetIO.RowSetter;
    use SheetIO.RowSetter.{RowSetter};
    use SheetIO.Sheet;
    use SheetIO.Sheet.{Sheet};

    use SheetIO.SheetWriter;
    use SheetIO.SheetWriter.{SheetWriter};
    use SheetIO.Cursors;

    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: SheetWriter[a], 
                    path: System.FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)

    /// Excel 2007 OOXML (.xlsx) format.
    /// Path is for a *new* file...
    pub def runXssf(ma: SheetWriter[a], 
                    path: System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{newXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let outpath     = toString(path);
        let* wb         = newXSSFWorkbook();
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
        let* _          = writeXSSFWorkbook(wb, outpath);
        let* _          = closeXSSFWorkbook(wb);
        Ok(ans)

    /// TODO - is call to `writeXSSFWorkbook` the right thing to do?
    pub def runTemplateXssf(ma: SheetWriter[a], 
                            template: System.FilePath, 
                            outPath: System.FilePath): Result[a, String] & Impure =
        use SheetIO/Internal/Prim.{openXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let* _          = System/File.copy(template, outPath) |> Result.mapErr(x -> System/Error.toString(x));
        let path1       = toString(outPath);
        let* wb         = openXSSFWorkbook(path1);
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());        
        let* _          = writeXSSFWorkbook(wb, path1);
        let* _          = closeXSSFWorkbook(wb);
        Ok(ans)

                    
    /// Excel 97 binary (.xls) format.
    pub def runHssf(ma: SheetWriter[a], path: System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{newHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};
        use Result.{flatMap};
        let path1       = toString(path);
        let* wb         = newHSSFWorkbook();
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
        let* _          = writeHSSFWorkbook(wb, path1);
        let* _          = closeHSSFWorkbook(wb);
        Ok(ans)

    pub def runTemplateHssf(ma: SheetWriter[a], 
                            template: System.FilePath, 
                            outPath: System.FilePath): Result[a, String] & Impure =
        use Result.{flatMap};
        use SheetIO/Internal/Prim.{openHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};        
        let* _          = System/File.copy(template, outPath) |> Result.mapErr(x -> System/Error.toString(x));
        let outPath1    = toString(outPath);
        let* wb         = openHSSFWorkbook(outPath1);
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
        let* _          = writeHSSFWorkbook(wb, outPath1);
        let* _          = closeHSSFWorkbook(wb);
        Ok(ans)

    def apply1(ma: SheetWriter[a], 
               wb: ##org.apache.poi.ss.usermodel.Workbook, 
               style: Option[CellStyle],
               s1: Cursors): Result[(Cursors, a), String] & Impure = 
        let SheetWriter(x1) = ma;
        x1(wb, style, s1)

    pub def askStyle(): SheetWriter[Option[CellStyle]] = 
        SheetWriter((_, style, s1) -> Ok((s1, style)) as & Impure)


    pub def return(x: a): SheetWriter[a] = 
        SheetWriter((_, _, s1) -> Ok((s1, x)) as & Impure)

    pub def throwError(msg: String): SheetWriter[a] =
        SheetWriter((_, _, _) -> Err(msg) as & Impure)


    pub def catchError(ma: SheetWriter[a], handle: String -> SheetWriter[a]): SheetWriter[a] =
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => Ok((s2, a))
            case Err(e) => apply1(handle(e), wb, style, s1)
        })

    pub def mapError(f: String -> String, ma: SheetWriter[a]): SheetWriter[a] =
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => Ok((s2, a))
            case Err(e) => Err(f(e))
        })

    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: SheetWriter[a]): SheetWriter[b] = 
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => { let ans = fn(a) as & Impure; Ok((s2, ans)) }
            case Err(e) => Err(e)
        })

    pub def ap(mf: SheetWriter[a -> b & ef], ma: SheetWriter[a]): SheetWriter[b] = 
        SheetWriter((wb, style, s1) -> match apply1(mf, wb, style, s1) { 
            case Ok((s2, f)) => match apply1(ma, wb, style, s2) { 
                case Ok((s3, a)) => { let ans = f(a) as & Impure; Ok((s3, ans)) }
                case Err(e) => Err(e)
            }
            case Err(e) => Err(e)
        })

    pub def seqLeft(ma: SheetWriter[a], mb: SheetWriter[b]): SheetWriter[a] = 
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => match apply1(mb, wb, style, s2) {
                case Ok((s3, _)) => Ok((s3, a))
                case Err(e) => Err(e)
            }
            case Err(e) => Err(e)
        })

    pub def seqRight(ma: SheetWriter[a], mb: SheetWriter[b]): SheetWriter[b] = 
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, _)) => apply1(mb, wb, style, s2)
            case Err(e) => Err(e)
        })


    pub def bind(ma: SheetWriter[a], fn: a -> SheetWriter[b]): SheetWriter[b] = 
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => apply1(fn(a), wb, style, s2)
            case Err(e) => Err(e)
        })

    pub def flatMap(fn: a -> SheetWriter[b], ma: SheetWriter[a]): SheetWriter[b] = bind(ma, fn)

    pub def kleisliLeft(f1: a -> SheetWriter[b], f2: b -> SheetWriter[c], x: a): SheetWriter[c] =
        SheetWriter((wb, style, s1) -> match apply1(f1(x), wb, style, s1) { 
            case Ok((s2, b)) => apply1(f2(b), wb, style, s2)
            case Err(e) => Err(e)
        })

    pub def kleisliRight(f1: b -> SheetWriter[c], f2: a -> SheetWriter[b], x: a): SheetWriter[c] =
        kleisliLeft(f2, f1, x)


    pub def guard(x: Bool): SheetWriter[Unit] = 
        match x { 
            case true => return()
            case false => throwError("guard")
        }

    pub def whenTrue(x: Bool, ma: SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case true => ma
            case false => return()
        }

    pub def whenFalse(x: Bool, ma: SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case true => return()
            case false => ma
        }

    pub def whenSome(x: Option[a], f: a -> SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case Some(a) => f(a)
            case None => return()
        }

    pub def whenNone(x: Option[a], ma: SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case Some(_) => return()
            case None => ma
        }



    def liftWorkbookPrim(fn: ##org.apache.poi.ss.usermodel.Workbook ~> Result[a, String]): SheetWriter[a] = 
        SheetWriter((wb, _, s1) -> {let ans = fn(wb); Result.map(x -> (s1, x), ans)})

    pub def liftAction(fn: Unit ~> a): SheetWriter[a] = 
        SheetWriter((_, _, s1)-> {let a = fn(); Ok((s1, a)) })

    pub def liftActionResult(fn: Unit ~> Result[a, String]): SheetWriter[a] = 
        SheetWriter((_, _, s1) -> {let ans = fn(); Result.map(x -> (s1, x), ans)})

    pub def liftOption(x: Option[a], errMsg: String): SheetWriter[a] = match x { 
        case Some(x1) => return(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, e], makeErrMsg: e -> String): SheetWriter[a] = match x { 
        case Ok(x1) => return(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetWriter[a] = match x { 
        case Success(x1) => return(x1)
        case Failure(es) => throwError(makeErrMsg(Nel.toList(es)))
    }

    pub def localCellStyle(style: CellStyle, mx: SheetWriter[a]): SheetWriter[a] = 
        SheetWriter((wb, _, s1) -> apply1(mx, wb, Some(style), s1))


    pub def mapM(mf: a -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] =
        SheetWriter((wb, style, s1) ->
            mapMHelper(wb, style, s1, mf, xs, e -> Err(e), (s2, ys) -> Ok(s2, ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   style: Option[CellStyle],
                   s1: Cursors,
                   mf: a -> SheetWriter[b], 
                   xs: List[a], 
                   fk: String -> Result[(Cursors, List[b]), String], 
                   sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), String]): Result[(Cursors, List[b]), String] & Impure =
        match xs {
            case Nil => sk(s1, Nil)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, x1)) => mapMHelper(wb, style, s2, mf, rs, fk, (s3, ks) -> sk(s3, x1 :: ks))
                }
            }
        }

    pub def mapX(mf: a -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] =
        SheetWriter((wb, style, s1) ->
            mapXHelper(wb, style, s1, mf, xs, e -> Err(e), s2 -> Ok((s2, ())))
        )

    def mapXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   style: Option[CellStyle],
                   s1: Cursors,
                   mf: a -> SheetWriter[b], 
                   xs: List[a], 
                   fk: String -> Result[(Cursors, Unit), String], 
                   sk: Cursors -> Result[(Cursors, Unit), String]): Result[(Cursors, Unit), String] & Impure =
        match xs {
            case Nil => sk(s1)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, _)) => mapXHelper(wb, style, s2, mf, rs, fk, sk)
                }
            }
        }

    pub def mapWithIndexM(mf: (a, Int32) -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] =
        SheetWriter((wb, style, s1) ->
            mapWithIndexMHelper(wb, style, s1, mf, 0, xs, e -> Err(e), (s2, ys) -> Ok((s2, ys)))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            style: Option[CellStyle],
                            s1: Cursors,
                            mf: (a, Int32) -> SheetWriter[b], 
                            ix : Int32, 
                            xs: List[a], 
                            fk: String -> Result[(Cursors, List[b]), String], 
                            sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), String]): Result[(Cursors, List[b]), String] & Impure =
        match xs {
            case Nil => sk(s1, Nil)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, x1)) => mapWithIndexMHelper(wb, style, s2, mf, ix + 1, rs, fk, (s3, ks) -> sk(s3, x1 :: ks))
                }
            }
        }

    pub def mapWithIndexX(mf: (a, Int32) -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] =
        SheetWriter((wb, style, s1) ->
            mapWithIndexXHelper(wb, style, s1, mf, 0, xs, e -> Err(e), s2 -> Ok((s2, ())))
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            style: Option[CellStyle],
                            s1: Cursors,
                            mf: (a, Int32) -> SheetWriter[b],
                            ix: Int32, 
                            xs: List[a], 
                            fk: String -> Result[(Cursors, Unit), String], 
                            sk: Cursors -> Result[(Cursors, Unit), String]): Result[(Cursors, Unit), String] & Impure =
        match xs {
            case Nil => sk(s1)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, _)) => mapWithIndexXHelper(wb, style, s2, mf, ix + 1, rs, fk, sk)
                }
            }
        }


    pub def sequence(xs: List[SheetWriter[a]]): SheetWriter[List[a]] =
        SheetWriter((wb, style, s1) ->
            sequenceHelper(wb, style, s1, xs, e -> Err(e), (s2, ks) -> Ok((s2, ks)))
        )

    ///
    /// Helper function for `sequence`.
    ///
    def sequenceHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                       style: Option[CellStyle],
                       s1: Cursors,
                       xs: List[SheetWriter[a]], 
                       fk: String -> Result[(Cursors, List[a]), String], 
                       sk: (Cursors, List[a]) -> Result[(Cursors, List[a]), String]): Result[(Cursors, List[a]), String] & Impure =
        match xs {
            case Nil => sk(s1, Nil)
            case mx :: rs => match apply1(mx, wb, style, s1) {
                case Err(e) => fk(e)
                case Ok((s2, x)) => sequenceHelper(wb, style, s2, rs, fk, (s3, ks) -> sk(s3, x :: ks))
            }
        }

    ///
    /// Returns `(v1 :: v2 :: ... v :: vn)` if each of `f(xs_i)` is successful.
    /// Otherwise fails with the first error encountered.
    ///
    /// `traverse` is a synonym for `mapM`.
    ///
    pub def traverse(f: a -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] = mapM(f, xs)

    ///
    /// `traverseX` is a synonym for `mapX`.
    ///
    pub def traverseX(f: a -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] = mapX(f, xs)

    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    pub def forM(xs: List[a], f: a -> SheetWriter[b]): SheetWriter[List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    pub def forX(xs: List[a], f: a -> SheetWriter[b]): SheetWriter[Unit] = mapX(f, xs)


    pub def foldM(mf: (b, a) -> SheetWriter[b], start: b, xs: List[a]): SheetWriter[b] =
        SheetWriter((wb, style, s1) ->
            foldMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac) -> Ok((s2, ac)))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    style: Option[CellStyle],
                    s1: Cursors,
                    mf: (b, a) -> SheetWriter[b], 
                    acc: b, 
                    xs: List[a], 
                    fk: String -> Result[(Cursors, b), String], 
                    sk: (Cursors, b) -> Result[(Cursors, b), String]): Result[(Cursors, b), String] & Impure =
        match xs {
            case Nil => sk(s1, acc)
            case x :: rs => {
                let ma = mf(acc, x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, acc1)) => foldMHelper(wb, style, s2, mf, acc1, rs, fk, sk)
                }
            }
        }

    pub def mapAccumM(mf: (ac, a) -> SheetWriter[(ac, b)], start: ac, xs: List[a]): SheetWriter[(ac, List[b])] =
        SheetWriter((wb, style, s1) ->
            mapAccumMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac, ks) -> Ok((s2, (ac, ks))))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        style: Option[CellStyle],
                        s1: Cursors,
                        mf: (ac, a) -> SheetWriter[(ac, b)],
                        ac: ac,
                        xs: List[a],
                        fk: String -> Result[(Cursors, (ac, List[b])), String],
                        sk: (Cursors, ac, List[b]) -> Result[(Cursors, (ac, List[b])), String]): Result[(Cursors, (ac, List[b])), String] & Impure =
        match xs {
            case Nil => sk(s1, ac, Nil)
            case x :: rs => {
                let ma = mf(ac, x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, (ac1, x1))) => mapAccumMHelper(wb, style, s2, mf, ac1, rs, fk, (s3, ac2, ks) -> sk(s3, ac2, x1 :: ks))
                }
            }
        }

    ///
    /// `mapAccumX` is a synonym for `foldM`.
    ///
    pub def mapAccumX(mf: (ac, a) -> SheetWriter[ac], start: ac, xs: List[a]): SheetWriter[ac] = foldM(mf, start, xs)

    /// Alias for `seqRight`
    pub def *>(p1: SheetWriter[a], p2: SheetWriter[b]): SheetWriter[b] = seqRight(p1, p2)

    /// Alias for `seqLeft`
    pub def <*(p1: SheetWriter[a], p2: SheetWriter[b]): SheetWriter[a] = seqLeft(p1, p2)

    // Haskell's <$>
    pub def <&>(f: a -> b, p2: SheetWriter[a]): SheetWriter[b] =
        map(f, p2)

    pub def <&(x: b, p2: SheetWriter[a]): SheetWriter[b] =
        p2 *> return(x)

    /// Alias for `bind`
    pub def >>=(ma: SheetWriter[a], k: a -> SheetWriter[b]): SheetWriter[b] = bind(ma, k)

    /// Alias for `flatMap` (flipped `bind`)
    pub def =<<(k: a -> SheetWriter[b], ma: SheetWriter[a]): SheetWriter[b] = flatMap(k, ma)


    // Haskell's ap <*>
    pub def <*>(mf: SheetWriter[a -> b], p2: SheetWriter[a]): SheetWriter[b] =
        ap(mf, p2)

    /// Alias for `kleisliLeft`
    pub def >=>(f1: a -> SheetWriter[b], f2: b -> SheetWriter[c]): a -> SheetWriter[c] = x -> kleisliLeft(f1, f2, x)

    /// Alias for `kleisliRight`
    pub def <=<(f1: b -> SheetWriter[c], f2: a -> SheetWriter[b]): a -> SheetWriter[c] = x -> kleisliRight(f1, f2, x)


    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2]): SheetWriter[u] =
        ap(map(f, o1), o2)


    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3]): SheetWriter[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3], 
                  o4: SheetWriter[t4]): SheetWriter[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3], 
                  o4: SheetWriter[t4], 
                  o5: SheetWriter[t5]): SheetWriter[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3],
                  o4: SheetWriter[t4], 
                  o5: SheetWriter[t5], 
                  o6: SheetWriter[t6]): SheetWriter[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3], 
                  o4: SheetWriter[t4], 
                  o5: SheetWriter[t5], 
                  o6: SheetWriter[t6], 
                  o7: SheetWriter[t7]): SheetWriter[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3], 
                  o4: SheetWriter[t4], 
                  o5: SheetWriter[t5], 
                  o6: SheetWriter[t6], 
                  o7: SheetWriter[t7], 
                  o8: SheetWriter[t8]): SheetWriter[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef, 
                  o1: SheetWriter[t1], 
                  o2: SheetWriter[t2], 
                  o3: SheetWriter[t3], 
                  o4: SheetWriter[t4], 
                  o5: SheetWriter[t5], 
                  o6: SheetWriter[t6], 
                  o7: SheetWriter[t7], 
                  o8: SheetWriter[t8], 
                  o9: SheetWriter[t9]): SheetWriter[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef, 
                   o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4], 
                   o5: SheetWriter[t5], 
                   o6: SheetWriter[t6], 
                   o7: SheetWriter[t7], 
                   o8: SheetWriter[t8], 
                   o9: SheetWriter[t9], 
                   o10: SheetWriter[t10]): SheetWriter[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)


    ///
    /// Returns a 2 element tuple with the results of applying the evaluators `o1` and `o2`.
    ///
    /// Returns an error if evaluating either `o1` or `o2` is an error.
    ///
    pub def tuple2(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2]): SheetWriter[(t1, t2)] =
        lift2((x,y) -> (x,y), o1, o2)

    ///
    /// Returns a 3 element tuple with the results of applying the evaluators `o1`, `o2` and `o3`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2` and `o3` is an error.
    ///
    pub def tuple3(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3]): SheetWriter[(t1, t2, t3)] =
        lift3((x1, x2, x3) -> (x1, x2, x3), o1, o2, o3)

    ///
    /// Returns a 4 element tuple with the results of applying the evaluators `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, `o3` and `o4` is an error.
    ///
    pub def tuple4(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4]): SheetWriter[(t1, t2, t3, t4)]  =
        lift4((x1, x2, x3, x4) -> (x1, x2, x3, x4), o1, o2, o3, o4)

    ///
    /// Returns a 5 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o5` is an error.
    ///
    pub def tuple5(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4], 
                   o5: SheetWriter[t5]): SheetWriter[(t1, t2, t3, t4, t5)]  =
        lift5((x1, x2, x3, x4, x5) -> (x1, x2, x3, x4, x5), o1, o2, o3, o4, o5)

    ///
    /// Returns a 6 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o6` is an error.
    ///
    pub def tuple6(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4], 
                   o5: SheetWriter[t5], 
                   o6: SheetWriter[t6]): SheetWriter[(t1, t2, t3, t4, t5, t6)]  =
        lift6((x1, x2, x3, x4, x5, x6) -> (x1, x2, x3, x4, x5, x6), o1, o2, o3, o4, o5, o6)

    ///
    /// Returns a 7 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o7` is an error.
    ///
    pub def tuple7(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4], 
                   o5: SheetWriter[t5], 
                   o6: SheetWriter[t6], 
                   o7: SheetWriter[t7]): SheetWriter[(t1, t2, t3, t4, t5, t6, t7)]  =
        lift7((x1, x2, x3, x4, x5, x6, x7) -> (x1, x2, x3, x4, x5, x6, x7), 
                o1, o2, o3, o4, o5, o6, o7)

    ///
    /// Returns a 8 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o8` is an error.
    ///
    pub def tuple8(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4], 
                   o5: SheetWriter[t5], 
                   o6: SheetWriter[t6], 
                   o7: SheetWriter[t7], 
                   o8: SheetWriter[t8]): SheetWriter[(t1, t2, t3, t4, t5, t6, t7, t8)]  =
        lift8((x1, x2, x3, x4, x5, x6, x7, x8) -> (x1, x2, x3, x4, x5, x6, x7, x8), 
                o1, o2, o3, o4, o5, o6, o7, o8)

    ///
    /// Returns a 9 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o9` is an error.
    ///
    pub def tuple9(o1: SheetWriter[t1], 
                   o2: SheetWriter[t2], 
                   o3: SheetWriter[t3], 
                   o4: SheetWriter[t4], 
                   o5: SheetWriter[t5], 
                   o6: SheetWriter[t6], 
                   o7: SheetWriter[t7], 
                   o8: SheetWriter[t8], 
                   o9: SheetWriter[t9]): SheetWriter[(t1, t2, t3, t4, t5, t6, t7, t8, t9)]  =
        lift9((x1, x2, x3, x4, x5, x6, x7, x8, x9) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9)

    ///
    /// Returns a 10 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o10` is an error.
    ///
    pub def tuple10(o1: SheetWriter[t1], 
                    o2: SheetWriter[t2], 
                    o3: SheetWriter[t3], 
                    o4: SheetWriter[t4], 
                    o5: SheetWriter[t5], 
                    o6: SheetWriter[t6], 
                    o7: SheetWriter[t7], 
                    o8: SheetWriter[t8], 
                    o9: SheetWriter[t9], 
                    o10: SheetWriter[t10]): SheetWriter[(t1, t2, t3, t4, t5, t6, t7, t8, t9, t10)]  =
        lift10((x1, x2, x3, x4, x5, x6, x7, x8, x9, x10) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9, x10), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9, o10)



    // SheetWriter

    def getCursorAndIncr(sheetName: String): SheetWriter[Int32] = 
        SheetWriter((_, _, s1) -> match Map.get(sheetName, s1) {
            case Some(ix)   => Ok((Map.adjust(x -> x + 1, sheetName, s1), ix)) as & Impure
            case None       => Ok((Map.insert(sheetName, 1, s1), 0)) as & Impure
        })

    

    pub def createSheet(name: String): SheetWriter[Sheet] =
        use SheetIO/Internal/Prim.{createSheet};
        liftWorkbookPrim(x1 -> Result.map(Sheet, x1.createSheet(name))) 

    pub def getSheetNamed(name: String): SheetWriter[Sheet] =
        use SheetIO/Internal/Prim.{getSheet};
        liftWorkbookPrim(x1 -> Result.map(Sheet, x1.getSheet(name))) 
    
    /// TEMP <<<
    pub def getSheetNames(): SheetWriter[List[String]] =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0,n);
        mapM(getSheetNameWb, ixs)      
    
    def getNumberOfSheets(): SheetWriter[Int32] =
        use SheetIO/Internal/Prim.{getNumberOfSheets};
        liftWorkbookPrim(x1 -> x1.getNumberOfSheets())

    def getSheetNameWb(ix: Int32): SheetWriter[String] =
        use SheetIO/Internal/Prim.{getSheetName};
        liftWorkbookPrim(x1 -> x1.getSheetName(ix))
    /// TEMP >>>

    pub def createDataFormat(): SheetWriter[DataFormat] =
        use SheetIO/Internal/Prim.{createDataFormat};
        liftWorkbookPrim(x1 -> Result.map(DataFormat, x1.createDataFormat()))

    pub def getFormat(x: DataFormat, format: String): SheetWriter[Int16] =
        use SheetIO/Internal/Prim.{getFormat};
        let DataFormat(x1) = x;
        liftActionResult(_ -> x1.getFormat(format))


    pub def setCursor!(x: Sheet, n: Int32): SheetWriter[Unit] =
        let* name = getSheetName(x);
        setCursorHelper(name, n)

    def getSheetName(x: Sheet): SheetWriter[String] =
        use SheetIO/Internal/Sheet.{getSheetName};
        let Sheet(x1) = x;
        liftActionResult(_ -> x1.getSheetName())

    def setCursorHelper(sheetName: String, n: Int32): SheetWriter[Unit] = 
        SheetWriter((_, _, s1) -> Ok((Map.insert(sheetName, n, s1), ())) as & Impure)


    pub def createCellStyle(mx: CellStyleSetter[a]): SheetWriter[CellStyle] =
        use SheetIO/Internal/Prim.{createCellStyle};
        let CellStyleSetter(mf) = mx;
        let* style              = liftWorkbookPrim(x1 -> Result.map(CellStyle, x1.createCellStyle()));
        let CellStyle(style1)   = style;
        let* _                  = liftActionResult(_ -> mf(style1));
        return(style)



    pub def createFont(mx: FontSetter[a]): SheetWriter[Font] =
        use SheetIO/Internal/Prim.{createFont};
        let FontSetter(mf)      = mx;
        let* font               = liftWorkbookPrim(x1 -> Result.map(Font, x1.createFont()));
        let Font(font1)         = font;
        let* _                  = liftActionResult(_ -> mf(font1));
        return(font)

    // createRow
    def createRow(x: Sheet, ix: Int32): SheetWriter[Row] =
        use SheetIO/Internal/Sheet.{createRow};
        let Sheet(x1) = x;
        liftActionResult(_ -> Result.map(Row, x1.createRow(ix)))

    pub def appendRow(x: Sheet, mx: RowSetter[a]): SheetWriter[Unit] =
        let RowSetter(mf) = mx;
        let* name       = getSheetName(x);
        let* ix         = getCursorAndIncr(name);
        let* row        = createRow(x, ix);
        let Row(row1)   = row;
        let* style      = askStyle();
        let* _          = liftActionResult(_ -> mf(row1, style));
        return()

    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allRowsWithHeaders(x: Sheet, headers: Array[String], rows: List[a], writer: a -> RowSetter[Unit]): SheetWriter[Unit] = 
        let* _      = appendRow(x, SheetIO/RowSetter.putHeaders!(headers));
        let* _      = mapX(r -> appendRow(x, writer(r)), rows);
        return()



}