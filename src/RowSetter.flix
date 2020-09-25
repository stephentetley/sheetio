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

use SheetIO.RowSetter;
use SheetIO.RowSetter.{RowSetter};



namespace SheetIO {

    pub opaque type RowSetter[a] = ##org.apache.poi.ss.usermodel.Row ~> Result[a, System.Error]

}

namespace SheetIO/RowSetter {

    pub def return(x: a): RowSetter[a] =
        RowSetter(_ -> Ok(x) as & Impure)

    def apply1(x: RowSetter[a], r: ##org.apache.poi.ss.usermodel.Row): Result[a, System.Error] & Impure =
        let RowSetter(f) = x;
        f(r)


    pub def mapError(f: System.Error -> System.Error, ma: RowSetter[a]): RowSetter[a] =
        RowSetter(r -> match apply1(ma, r) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => Err(f(e))
        })


    pub def map(f: a -> b & ef, x: RowSetter[a]): RowSetter[b] =
        RowSetter(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: RowSetter[a -> b & ef], ma: RowSetter[a]): RowSetter[b] =
        RowSetter(r ->
            match apply1(mf, r) {
                case Ok(f) => {
                    match apply1(ma, r) {
                        case Ok(a) => Ok(f(a) as & Impure)
                        case Err(e2) => Err(e2)
                    }
                }
                case Err(e1) => Err(e1)
            }
        )

    pub def bind(x: RowSetter[a], k: a -> RowSetter[b]): RowSetter[b] =
        RowSetter(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a); apply1(ans, r) }
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(k: a -> RowSetter[b], x: RowSetter[a]): RowSetter[b] = bind(x, k)

    def liftPrim(update: ##org.apache.poi.ss.usermodel.Row ~> Result[a, System.Error]): RowSetter[a] =
        RowSetter(r -> update(r))



    // pub def putNull!(ix: Int32): RowSetter[Unit] =
    //     use FactIO/Internal//PreparedStatement.{sql_type_null};
    //     liftPrim(x -> FactIO/Internal//PreparedStatement.setNull!(x, ix + 1, sql_type_null()))

    pub def putString!(ix: Int32, value: String): RowSetter[Unit] =
        use Result.flatMap;
        let step = x -> {
            let* cell   = SheetIO/Internal/Row.createCell(x, ix);
            let* _      = SheetIO/Internal/Cell.setStringCellValue(cell, value);
            Ok()
        };
        liftPrim(step)

    pub def putBool!(ix: Int32, value: Bool): RowSetter[Unit] =
        use Result.flatMap;
        let step = x -> {
            let* cell   = SheetIO/Internal/Row.createCell(x, ix);
            let* _      = SheetIO/Internal/Cell.setBooleanCellValue(cell, value);
            Ok()
        };
        liftPrim(step)

    def putNumeric!(ix: Int32, errMsg: String, value: Float64): RowSetter[Unit] =
        use Result.flatMap;
        let step = x -> {
            let* cell   = SheetIO/Internal/Row.createCell(x, ix);
            let* _      = SheetIO/Internal/Cell.setNumericCellValue(cell, value);
            Ok()
        };
        mapError(_ -> Generic(errMsg), liftPrim(step))

    pub def putInt8!(ix: Int32, n: Int8): RowSetter[Unit] =
        putNumeric!(ix, "putInt8!", Int8.toFloat64(n))

    pub def putInt16!(ix: Int32, n: Int16): RowSetter[Unit] =
        putNumeric!(ix, "putInt16!", Int16.toFloat64(n))

    pub def putInt32!(ix: Int32, n: Int32): RowSetter[Unit] =
        putNumeric!(ix, "putInt32!", Int32.toFloat64(n))

    pub def putInt64!(ix: Int32, n: Int64): RowSetter[Unit] =
        putNumeric!(ix, "putInt64!", Int64.toFloat64(n))

    pub def putFloat32!(ix: Int32, n: Float32): RowSetter[Unit] =
        putNumeric!(ix, "putFloat32!", Float32.toFloat64(n))

    pub def putFloat64!(ix: Int32, n: Float64): RowSetter[Unit] =
        putNumeric!(ix, "putFloat64!", n)



    // pub def putOptionString!(ix: Int32, x: Option[String]): RowSetter[Unit] = match x {
    //     case Some(s) => putString!(ix, s)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionBool!(ix: Int32, x: Option[Bool]): RowSetter[Unit] = match x {
    //     case Some(b) => putBool!(ix, b)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionInt8!(ix: Int32, x: Option[Int8]): RowSetter[Unit] = match x {
    //     case Some(n) => putInt8!(ix, n)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionInt16!(ix: Int32, x: Option[Int16]): RowSetter[Unit] = match x {
    //     case Some(n) => putInt16!(ix, n)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionInt32!(ix: Int32, x: Option[Int32]): RowSetter[Unit] = match x {
    //     case Some(n) => putInt32!(ix, n)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionInt64!(ix: Int32, x: Option[Int64]): RowSetter[Unit] = match x {
    //     case Some(n) => putInt64!(ix, n)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionFloat32!(ix: Int32, x: Option[Float32]): RowSetter[Unit] = match x {
    //     case Some(n) => putFloat32!(ix, n)
    //     case None => putNull!(ix)
    // }

    // pub def putOptionFloat64!(ix: Int32, x: Option[Float64]): RowSetter[Unit] = match x {
    //     case Some(n) => putFloat64!(ix, n)
    //     case None => putNull!(ix)
    // }


    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, o1: RowSetter[t1], o2: RowSetter[t2]): RowSetter[u] =
        ap(map(f, o1), o2)

    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3]): RowSetter[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4]): RowSetter[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5]): RowSetter[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5],
                  o6: RowSetter[t6]): RowSetter[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5],
                  o6: RowSetter[t6],
                  o7: RowSetter[t7]): RowSetter[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef,
                      o1: RowSetter[t1],
                      o2: RowSetter[t2],
                      o3: RowSetter[t3],
                      o4: RowSetter[t4],
                      o5: RowSetter[t5],
                      o6: RowSetter[t6],
                      o7: RowSetter[t7],
                      o8: RowSetter[t8]): RowSetter[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef,
                  o1: RowSetter[t1],
                  o2: RowSetter[t2],
                  o3: RowSetter[t3],
                  o4: RowSetter[t4],
                  o5: RowSetter[t5],
                  o6: RowSetter[t6],
                  o7: RowSetter[t7],
                  o8: RowSetter[t8],
                  o9: RowSetter[t9]): RowSetter[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef,
                   o1: RowSetter[t1],
                   o2: RowSetter[t2],
                   o3: RowSetter[t3],
                   o4: RowSetter[t4],
                   o5: RowSetter[t5],
                   o6: RowSetter[t6],
                   o7: RowSetter[t7],
                   o8: RowSetter[t8],
                   o9: RowSetter[t9],
                   o10: RowSetter[t10]): RowSetter[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)


}

