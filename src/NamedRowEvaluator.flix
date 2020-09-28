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

use System.Error.{Generic, IllegalArgumentException};

use SheetIO.NamedRowEvaluator;
use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};

use SheetIO.CellType;


namespace SheetIO {

    pub opaque type NamedRowEvaluator[a] = (Map[String, Int32], ##org.apache.poi.ss.usermodel.Row) ~> Result[a, System.Error]

}

namespace SheetIO/NamedRowEvaluator {

    pub def makeHeaders(names: Array[String]): Map[String, Int32] & Impure =
        let step = (acc, x) -> { let (i, mx) = acc; (i + 1, Map.insert(x, i, mx)) };
        Array.foldLeft(step, (0, Map.empty()), names) |> snd


    pub def return(x: a): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _) -> Ok(x) as & Impure)

    def apply1(x: NamedRowEvaluator[a], 
               ixs: Map[String, Int32], 
               row: ##org.apache.poi.ss.usermodel.Row): Result[a, System.Error] & Impure =
        let NamedRowEvaluator(f) = x;
        f(ixs, row)

    pub def throwError(msg: String): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _) -> Err(System/Error.Generic(msg)) as & Impure)


    pub def map(f: a -> b & ef, x: NamedRowEvaluator[a]): NamedRowEvaluator[b] =
        NamedRowEvaluator((ixs, r) ->
            match apply1(x, ixs, r) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: NamedRowEvaluator[a -> b & ef], ma: NamedRowEvaluator[a]): NamedRowEvaluator[b] =
        NamedRowEvaluator((ixs, r) ->
            match apply1(mf, ixs, r) {
                case Ok(f) => {
                    match apply1(ma, ixs, r) {
                        case Ok(a) => Ok(f(a) as & Impure)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        )
    
    pub def seqLeft(ma: NamedRowEvaluator[a], mb: NamedRowEvaluator[b]): NamedRowEvaluator[a] = 
        NamedRowEvaluator((ixs, r) ->
            match apply1(ma, ixs, r) { 
                case Ok(a) => match apply1(mb, ixs, r) {
                    case Ok(_) => Ok(a)
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqRight(ma: NamedRowEvaluator[a], mb: NamedRowEvaluator[b]): NamedRowEvaluator[b] = 
        NamedRowEvaluator((ixs, r) ->
            match apply1(ma, ixs, r) { 
                case Ok(_) => apply1(mb, ixs, r)
                case Err(e) => Err(e)
            }
        )

    pub def bind(x: NamedRowEvaluator[a], k: a -> NamedRowEvaluator[b]): NamedRowEvaluator[b] =
        NamedRowEvaluator((ixs, r) ->
            match apply1(x, ixs, r) {
                case Ok(a) => { let ans = k(a); apply1(ans, ixs, r) }
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(k: a -> NamedRowEvaluator[b], x: NamedRowEvaluator[a]): NamedRowEvaluator[b] = bind(x, k)

    pub def kleisliLeft(f1: a -> NamedRowEvaluator[b], f2: b -> NamedRowEvaluator[c], x: a): NamedRowEvaluator[c] =
        NamedRowEvaluator((ixs, r) -> 
            match apply1(f1(x), ixs, r) { 
                case Ok(b) => apply1(f2(b), ixs, r)
                case Err(e) => Err(e)
            }
        )

    pub def kleisliRight(f1: b -> NamedRowEvaluator[c], f2: a -> NamedRowEvaluator[b], x: a): NamedRowEvaluator[c] =
        kleisliLeft(f2, f1, x)

    pub def optional(x: NamedRowEvaluator[a]): NamedRowEvaluator[Option[a]] =
        NamedRowEvaluator((ixs, r) ->
            match apply1(x, ixs, r) {
                case Ok(a) => Ok(Some(a))
                case Err(_) => Ok(None)
            }
        )


    def liftPrim(ask: ##org.apache.poi.ss.usermodel.Row ~> Result[a, System.Error]): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, r) -> ask(r))


    pub def liftAction(fn: Unit ~> a): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _) -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit ~> Result[a, System.Error]): NamedRowEvaluator[a] =
        NamedRowEvaluator((_, _) -> fn())



    // // Note - the primitive setX! functions index from 1, we add 1 to calls
    // // so FactIO's API indexes from 0.
    
    def getIndex(label: String): NamedRowEvaluator[Int32] =
        NamedRowEvaluator((ixs, _) -> match Map.get(label, ixs) {
            case None => Err(IllegalArgumentException) as & Impure
            case Some(ix) => Ok(ix) as & Impure
        })


    def withCell(row: ##org.apache.poi.ss.usermodel.Row,
                 ix: Int32, 
                 f: ##org.apache.poi.ss.usermodel.Cell ~> Result[a, System.Error]): Result[a, System.Error] & Impure = 
        use Result.{flatMap};
        let* cell = SheetIO/Internal/Row.getCell(row, ix);
        f(cell)
    
    pub def isNullNamed(label: String): NamedRowEvaluator[Bool] = 
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, ans -> Ok(Object.isNull(ans) as & Impure)))

    pub def getNullNamed(label: String): NamedRowEvaluator[Unit] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, ans -> if (Object.isNull(ans)) (Ok() as & Impure) else Err(Generic("getNull"))))


    pub def getCellTypeNamed(label: String): NamedRowEvaluator[CellType] =
        let* ix     = getIndex(label);
        map(SheetIO/CellType.unmarshal, liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getCellType)))

    
    pub def getStringNamed(label: String): NamedRowEvaluator[String] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getStringCellValue))

    pub def getStringNamedWithDefault(label: String, d: String): NamedRowEvaluator[String] =
        let* x = isNullNamed(label);
        if (x) return(d) else getStringNamed(label)

    pub def getBoolNamed(label: String): NamedRowEvaluator[Bool] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getBooleanCellValue))
    

    pub def getNumericNamed(label: String): NamedRowEvaluator[Float64] =
        let* ix     = getIndex(label);
        liftPrim(x -> withCell(x, ix, SheetIO/Internal/Cell.getNumericCellValue))


    def getNumNamed(label: String, conv: Float64 -> Option[a], errMsg: String): NamedRowEvaluator[a] =
        let* n = getNumericNamed(label);
        match conv(n) {
            case Some(i) => return(i)
            case None => throwError(errMsg)
        }


    pub def getInt8Named(label: String): NamedRowEvaluator[Int8] =
        getNumNamed(label, Float64.tryToInt8, "getInt8")

    pub def getInt16Named(label: String): NamedRowEvaluator[Int16] =
        getNumNamed(label, Float64.tryToInt16, "getInt16")

    pub def getInt32Named(label: String): NamedRowEvaluator[Int32] =
        getNumNamed(label, Float64.tryToInt32, "getInt32")

    pub def getInt64Named(label: String): NamedRowEvaluator[Int64] =
        getNumNamed(label, Float64.tryToInt64, "getInt64")

    pub def getFloat32At(label: String): NamedRowEvaluator[Float32] =
        getNumNamed(label, Float64.tryToFloat32, "getFloat32")
    
    pub def getFloat64Named(label: String): NamedRowEvaluator[Float64] = 
        getNumNamed(label, Some, "getFloat32")



    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2]): NamedRowEvaluator[u] =
        ap(map(f, o1), o2)

    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3]): NamedRowEvaluator[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3], 
                  o4: NamedRowEvaluator[t4]): NamedRowEvaluator[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3], 
                  o4: NamedRowEvaluator[t4], 
                  o5: NamedRowEvaluator[t5]): NamedRowEvaluator[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3],
                  o4: NamedRowEvaluator[t4], 
                  o5: NamedRowEvaluator[t5], 
                  o6: NamedRowEvaluator[t6]): NamedRowEvaluator[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3], 
                  o4: NamedRowEvaluator[t4], 
                  o5: NamedRowEvaluator[t5], 
                  o6: NamedRowEvaluator[t6], 
                  o7: NamedRowEvaluator[t7]): NamedRowEvaluator[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3], 
                  o4: NamedRowEvaluator[t4], 
                  o5: NamedRowEvaluator[t5], 
                  o6: NamedRowEvaluator[t6], 
                  o7: NamedRowEvaluator[t7], 
                  o8: NamedRowEvaluator[t8]): NamedRowEvaluator[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef, 
                  o1: NamedRowEvaluator[t1], 
                  o2: NamedRowEvaluator[t2], 
                  o3: NamedRowEvaluator[t3], 
                  o4: NamedRowEvaluator[t4], 
                  o5: NamedRowEvaluator[t5], 
                  o6: NamedRowEvaluator[t6], 
                  o7: NamedRowEvaluator[t7], 
                  o8: NamedRowEvaluator[t8], 
                  o9: NamedRowEvaluator[t9]): NamedRowEvaluator[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef, 
                   o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4], 
                   o5: NamedRowEvaluator[t5], 
                   o6: NamedRowEvaluator[t6], 
                   o7: NamedRowEvaluator[t7], 
                   o8: NamedRowEvaluator[t8], 
                   o9: NamedRowEvaluator[t9], 
                   o10: NamedRowEvaluator[t10]): NamedRowEvaluator[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)


    ///
    /// Returns a 2 element tuple with the results of applying the evaluators `o1` and `o2`.
    ///
    /// Returns an error if evaluating either `o1` or `o2` is an error.
    ///
    pub def tuple2(o1: NamedRowEvaluator[t1], o2: NamedRowEvaluator[t2]): NamedRowEvaluator[(t1, t2)] =
        lift2((x,y) -> (x,y), o1, o2)

    ///
    /// Returns a 3 element tuple with the results of applying the evaluators `o1`, `o2` and `o3`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2` and `o3` is an error.
    ///
    pub def tuple3(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3]): NamedRowEvaluator[(t1, t2, t3)] =
        lift3((x1, x2, x3) -> (x1, x2, x3), o1, o2, o3)

    ///
    /// Returns a 4 element tuple with the results of applying the evaluators `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, `o3` and `o4` is an error.
    ///
    pub def tuple4(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4]): NamedRowEvaluator[(t1, t2, t3, t4)]  =
        lift4((x1, x2, x3, x4) -> (x1, x2, x3, x4), o1, o2, o3, o4)

    ///
    /// Returns a 5 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o5` is an error.
    ///
    pub def tuple5(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4], 
                   o5: NamedRowEvaluator[t5]): NamedRowEvaluator[(t1, t2, t3, t4, t5)]  =
        lift5((x1, x2, x3, x4, x5) -> (x1, x2, x3, x4, x5), o1, o2, o3, o4, o5)

    ///
    /// Returns a 6 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o6` is an error.
    ///
    pub def tuple6(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4], 
                   o5: NamedRowEvaluator[t5], 
                   o6: NamedRowEvaluator[t6]): NamedRowEvaluator[(t1, t2, t3, t4, t5, t6)]  =
        lift6((x1, x2, x3, x4, x5, x6) -> (x1, x2, x3, x4, x5, x6), o1, o2, o3, o4, o5, o6)

    ///
    /// Returns a 7 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o7` is an error.
    ///
    pub def tuple7(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4], 
                   o5: NamedRowEvaluator[t5], 
                   o6: NamedRowEvaluator[t6], 
                   o7: NamedRowEvaluator[t7]): NamedRowEvaluator[(t1, t2, t3, t4, t5, t6, t7)]  =
        lift7((x1, x2, x3, x4, x5, x6, x7) -> (x1, x2, x3, x4, x5, x6, x7), 
                o1, o2, o3, o4, o5, o6, o7)

    ///
    /// Returns a 8 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o8` is an error.
    ///
    pub def tuple8(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4], 
                   o5: NamedRowEvaluator[t5], 
                   o6: NamedRowEvaluator[t6], 
                   o7: NamedRowEvaluator[t7], 
                   o8: NamedRowEvaluator[t8]): NamedRowEvaluator[(t1, t2, t3, t4, t5, t6, t7, t8)]  =
        lift8((x1, x2, x3, x4, x5, x6, x7, x8) -> (x1, x2, x3, x4, x5, x6, x7, x8), 
                o1, o2, o3, o4, o5, o6, o7, o8)

    ///
    /// Returns a 9 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o9` is an error.
    ///
    pub def tuple9(o1: NamedRowEvaluator[t1], 
                   o2: NamedRowEvaluator[t2], 
                   o3: NamedRowEvaluator[t3], 
                   o4: NamedRowEvaluator[t4], 
                   o5: NamedRowEvaluator[t5], 
                   o6: NamedRowEvaluator[t6], 
                   o7: NamedRowEvaluator[t7], 
                   o8: NamedRowEvaluator[t8], 
                   o9: NamedRowEvaluator[t9]): NamedRowEvaluator[(t1, t2, t3, t4, t5, t6, t7, t8, t9)]  =
        lift9((x1, x2, x3, x4, x5, x6, x7, x8, x9) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9)

    ///
    /// Returns a 10 element tuple with the results of applying the evaluators `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if evaluating any of `o1`, `o2`, ... `o10` is an error.
    ///
    pub def tuple10(o1: NamedRowEvaluator[t1], 
                    o2: NamedRowEvaluator[t2], 
                    o3: NamedRowEvaluator[t3], 
                    o4: NamedRowEvaluator[t4], 
                    o5: NamedRowEvaluator[t5], 
                    o6: NamedRowEvaluator[t6], 
                    o7: NamedRowEvaluator[t7], 
                    o8: NamedRowEvaluator[t8], 
                    o9: NamedRowEvaluator[t9], 
                    o10: NamedRowEvaluator[t10]): NamedRowEvaluator[(t1, t2, t3, t4, t5, t6, t7, t8, t9, t10)]  =
        lift10((x1, x2, x3, x4, x5, x6, x7, x8, x9, x10) -> (x1, x2, x3, x4, x5, x6, x7, x8, x9, x10), 
                o1, o2, o3, o4, o5, o6, o7, o8, o9, o10)

    /// Alias for `seqRight`
    pub def *>(p1: NamedRowEvaluator[a], p2: NamedRowEvaluator[b]): NamedRowEvaluator[b] = seqRight(p1, p2)

    /// Alias for `seqLeft`
    pub def <*(p1: NamedRowEvaluator[a], p2: NamedRowEvaluator[b]): NamedRowEvaluator[a] = seqLeft(p1, p2)

    // Haskell's <$>
    pub def <&>(f: a -> b, p2: NamedRowEvaluator[a]): NamedRowEvaluator[b] =
        map(f, p2)

    pub def <&(x: b, p2: NamedRowEvaluator[a]): NamedRowEvaluator[b] =
        p2 *> return(x)

    /// Alias for `bind`
    pub def >>=(ma: NamedRowEvaluator[a], k: a -> NamedRowEvaluator[b]): NamedRowEvaluator[b] = 
        bind(ma, k)

    /// Alias for `flatMap` (flipped `bind`)
    pub def =<<(k: a -> NamedRowEvaluator[b], ma: NamedRowEvaluator[a]): NamedRowEvaluator[b] = 
        flatMap(k, ma)


    // Haskell's ap <*>
    pub def <*>(mf: NamedRowEvaluator[a -> b], p2: NamedRowEvaluator[a]): NamedRowEvaluator[b] =
        ap(mf, p2)

    /// Alias for `kleisliLeft`
    pub def >=>(f1: a -> NamedRowEvaluator[b], f2: b -> NamedRowEvaluator[c]): a -> NamedRowEvaluator[c] = 
        x -> kleisliLeft(f1, f2, x)

    /// Alias for `kleisliRight`
    pub def <=<(f1: b -> NamedRowEvaluator[c], f2: a -> NamedRowEvaluator[b]): a -> NamedRowEvaluator[c] = 
        x -> kleisliRight(f1, f2, x)


}

