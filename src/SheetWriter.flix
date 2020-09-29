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


use System.Error.{Generic};

use SheetIO.Row;
use SheetIO.Row.{Row};
use SheetIO.RowSetter;
use SheetIO.RowSetter.{RowSetter};
use SheetIO.Sheet;
use SheetIO.Sheet.{Sheet};

use SheetIO.SheetWriter;
use SheetIO.SheetWriter.{SheetWriter};
use SheetIO.Cursors;

namespace SheetIO {

   
    pub opaque type SheetWriter[a] = ##org.apache.poi.ss.usermodel.Workbook -> Cursors ~> Result[(Cursors, a), System.Error]

    type alias Cursors = Map[String, Int32]

}

namespace SheetIO/SheetWriter {

    
    pub def runXlsx(ma: SheetWriter[a], path: System.FilePath): Result[a, System.Error] & Impure = 
        use SheetIO/Internal/Prim.{newXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let outpath     = System/FilePath.toString(path);
        let* wb         = newXSSFWorkbook();
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, Map.empty());
        let* _          = writeXSSFWorkbook(wb, outpath);
        let* _          = closeXSSFWorkbook(wb);
        Ok(ans)

    def apply1(ma: SheetWriter[a], 
               wb: ##org.apache.poi.ss.usermodel.Workbook, 
               st: Cursors): Result[(Cursors, a), System.Error] & Impure = 
        let SheetWriter(x1) = ma;
        x1(wb, st)

    pub def return(x: a): SheetWriter[a] = 
        SheetWriter((_, st) -> Ok((st, x)) as & Impure)

    pub def throwError(msg: String): SheetWriter[a] =
        SheetWriter((_, _) -> Err(System/Error.Generic(msg)) as & Impure)


    pub def catchError(ma: SheetWriter[a], handle: System.Error -> SheetWriter[a]): SheetWriter[a] =
        SheetWriter((wb, st) -> match apply1(ma, wb, st) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => apply1(handle(e), wb, st)
        })

    pub def mapError(f: System.Error -> System.Error, ma: SheetWriter[a]): SheetWriter[a] =
        SheetWriter((wb, st) -> match apply1(ma, wb, st) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => Err(f(e))
        })

    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: SheetWriter[a]): SheetWriter[b] = 
        SheetWriter((wb, st) -> 
            match apply1(ma, wb, st) { 
                case Ok((st1, a)) => { let ans = fn(a) as & Impure; Ok((st1, ans)) }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: SheetWriter[a -> b & ef], ma: SheetWriter[a]): SheetWriter[b] = 
        SheetWriter((wb, st) -> 
            match apply1(mf, wb, st) { 
                case Ok((st1, f)) => match apply1(ma, wb, st1) { 
                    case Ok((st2, a)) => { let ans = f(a) as & Impure; Ok((st2, ans)) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqLeft(ma: SheetWriter[a], mb: SheetWriter[b]): SheetWriter[a] = 
        SheetWriter((wb, st) -> 
            match apply1(ma, wb, st) { 
                case Ok((st1, a)) => match apply1(mb, wb, st1) {
                    case Ok((st2, _)) => Ok((st2, a))
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqRight(ma: SheetWriter[a], mb: SheetWriter[b]): SheetWriter[b] = 
        SheetWriter((wb, st) -> 
            match apply1(ma, wb, st) { 
                case Ok((st1, _)) => apply1(mb, wb, st1)
                case Err(e) => Err(e)
            }
        )


    pub def bind(ma: SheetWriter[a], fn: a -> SheetWriter[b]): SheetWriter[b] = 
        SheetWriter((wb, st) -> 
            match apply1(ma, wb, st) { 
                case Ok((st1, a)) => apply1(fn(a), wb, st1)
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(fn: a -> SheetWriter[b], ma: SheetWriter[a]): SheetWriter[b] = bind(ma, fn)

    pub def kleisliLeft(f1: a -> SheetWriter[b], f2: b -> SheetWriter[c], x: a): SheetWriter[c] =
        SheetWriter((wb, st) -> 
            match apply1(f1(x), wb, st) { 
                case Ok((st1, b)) => apply1(f2(b), wb, st1)
                case Err(e) => Err(e)
            }
        )

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



    def liftWorkbookPrim(fn: ##org.apache.poi.ss.usermodel.Workbook ~> Result[a, System.Error]): SheetWriter[a] = 
        SheetWriter((wb, st) -> { let ans = fn(wb); Result.map(x -> (st, x), ans)})

    pub def liftAction(fn: Unit ~> a): SheetWriter[a] = 
        SheetWriter((_, st)-> { let a = fn(); Ok((st, a)) })

    pub def liftActionResult(fn: Unit ~> Result[a, System.Error]): SheetWriter[a] = 
        SheetWriter((_, st) -> { let ans = fn(); Result.map(x -> (st, x), ans)})

    
    
    pub def mapM(mf: a -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] =
        SheetWriter((wb, st) ->
            mapMHelper(wb, st, mf, xs, e -> Err(e), (st1, ys) -> Ok(st1, ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   st: Cursors,
                   mf: a -> SheetWriter[b], 
                   xs: List[a], 
                   fk: System.Error -> Result[(Cursors, List[b]), System.Error], 
                   sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), System.Error]): Result[(Cursors, List[b]), System.Error] & Impure =
        match xs {
            case Nil => sk(st, Nil)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb, st) {
                    case Err(e) => fk(e)
                    case Ok((st1, x1)) => mapMHelper(wb, st1, mf, rs, fk, (st2, ks) -> sk(st2, x1 :: ks))
                }
            }
        }

    pub def mapX(mf: a -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] =
        SheetWriter((wb, st) ->
            mapXHelper(wb, st, mf, xs, e -> Err(e), st1 -> Ok((st1, ())))
        )

    def mapXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   st: Cursors,
                   mf: a -> SheetWriter[b], 
                   xs: List[a], 
                   fk: System.Error -> Result[(Cursors, Unit), System.Error], 
                   sk: Cursors -> Result[(Cursors, Unit), System.Error]): Result[(Cursors, Unit), System.Error] & Impure =
        match xs {
            case Nil => sk(st)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb, st) {
                    case Err(e) => fk(e)
                    case Ok((st1, _)) => mapXHelper(wb, st1, mf, rs, fk, sk)
                }
            }
        }

    pub def mapWithIndexM(mf: (a, Int32) -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] =
        SheetWriter((wb, st) ->
            mapWithIndexMHelper(wb, st, mf, 0, xs, e -> Err(e), (st1, ys) -> Ok((st1, ys)))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            st: Cursors,
                            mf: (a, Int32) -> SheetWriter[b], 
                            ix : Int32, 
                            xs: List[a], 
                            fk: System.Error -> Result[(Cursors, List[b]), System.Error], 
                            sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), System.Error]): Result[(Cursors, List[b]), System.Error] & Impure =
        match xs {
            case Nil => sk(st, Nil)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb, st) {
                    case Err(e) => fk(e)
                    case Ok((st1, x1)) => mapWithIndexMHelper(wb, st1, mf, ix + 1, rs, fk, (st2, ks) -> sk(st2, x1 :: ks))
                }
            }
        }

    pub def mapWithIndexX(mf: (a, Int32) -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] =
        SheetWriter((wb, st) ->
            mapWithIndexXHelper(wb, st, mf, 0, xs, e -> Err(e), st1 -> Ok((st1, ())))
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            st: Cursors,
                            mf: (a, Int32) -> SheetWriter[b],
                            ix: Int32, 
                            xs: List[a], 
                            fk: System.Error -> Result[(Cursors, Unit), System.Error], 
                            sk: Cursors -> Result[(Cursors, Unit), System.Error]): Result[(Cursors, Unit), System.Error] & Impure =
        match xs {
            case Nil => sk(st)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb, st) {
                    case Err(e) => fk(e)
                    case Ok((st1, _)) => mapWithIndexXHelper(wb, st1, mf, ix + 1, rs, fk, sk)
                }
            }
        }


    pub def sequence(xs: List[SheetWriter[a]]): SheetWriter[List[a]] =
        SheetWriter((wb, st) ->
            sequenceHelper(wb, st, xs, e -> Err(e), (st1, ks) -> Ok((st1, ks)))
        )

    ///
    /// Helper function for `sequence`.
    ///
    def sequenceHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                       st: Cursors,
                       xs: List[SheetWriter[a]], 
                       fk: System.Error -> Result[(Cursors, List[a]), System.Error], 
                       sk: (Cursors, List[a]) -> Result[(Cursors, List[a]), System.Error]): Result[(Cursors, List[a]), System.Error] & Impure =
        match xs {
            case Nil => sk(st, Nil)
            case mx :: rs => match apply1(mx, wb, st) {
                case Err(e) => fk(e)
                case Ok((st1, x)) => sequenceHelper(wb, st1, rs, fk, (st2, ks) -> sk(st2, x :: ks))
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
        SheetWriter((wb, st) ->
            foldMHelper(wb, st, mf, start, xs, e -> Err(e), (st1, ac) -> Ok((st1, ac)))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    st: Cursors,
                    mf: (b, a) -> SheetWriter[b], 
                    acc: b, 
                    xs: List[a], 
                    fk: System.Error -> Result[(Cursors, b), System.Error], 
                    sk: (Cursors, b) -> Result[(Cursors, b), System.Error]): Result[(Cursors, b), System.Error] & Impure =
        match xs {
            case Nil => sk(st, acc)
            case x :: rs => {
                let ma = mf(acc, x);
                match apply1(ma, wb, st) {
                    case Err(e) => fk(e)
                    case Ok((st1, acc1)) => foldMHelper(wb, st1, mf, acc1, rs, fk, sk)
                }
            }
        }

    pub def mapAccumM(mf: (ac, a) -> SheetWriter[(ac, b)], start: ac, xs: List[a]): SheetWriter[(ac, List[b])] =
        SheetWriter((wb, st) ->
            mapAccumMHelper(wb, st, mf, start, xs, e -> Err(e), (st1, ac, ks) -> Ok((st1, (ac, ks))))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        st: Cursors,
                        mf: (ac, a) -> SheetWriter[(ac, b)],
                        ac: ac,
                        xs: List[a],
                        fk: System.Error -> Result[(Cursors, (ac, List[b])), System.Error],
                        sk: (Cursors, ac, List[b]) -> Result[(Cursors, (ac, List[b])), System.Error]): Result[(Cursors, (ac, List[b])), System.Error] & Impure =
        match xs {
            case Nil => sk(st, ac, Nil)
            case x :: rs => {
                let ma = mf(ac, x);
                match apply1(ma, wb, st) {
                    case Err(e) => fk(e)
                    case Ok((st1, (ac1, x1))) => mapAccumMHelper(wb, st1, mf, ac1, rs, fk, (st2, ac2, ks) -> sk(st2, ac2, x1 :: ks))
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
        SheetWriter((_, st) -> match Map.get(sheetName, st) {
            case Some(ix)   => Ok((Map.adjust(x -> x + 1, sheetName, st), ix)) as & Impure
            case None       => Ok((Map.insert(sheetName, 1, st), 0)) as & Impure
        })

    pub def createSheet(name: String): SheetWriter[Sheet] =
        use SheetIO/Internal/Prim.{createSheet};
        liftWorkbookPrim(x1 -> Result.map(Sheet, x1.createSheet(name))) 

    def getSheetName(x: Sheet): SheetWriter[String] =
        use SheetIO/Internal/Sheet.{getSheetName};
        let Sheet(x1) = x;
        liftActionResult(_ -> x1.getSheetName())


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
        let* _          = liftActionResult(_ -> mf(row1));
        return()

    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allRowsWithHeaders(x: Sheet, headers: Array[String], rows: List[a], writer: a -> RowSetter[Unit]): SheetWriter[Unit] = 
        let* _      = appendRow(x, SheetIO/RowSetter.putHeaders!(headers));
        let* _      = mapX(r -> appendRow(x, writer(r)), rows);
        return()



}