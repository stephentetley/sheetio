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

    pub opaque type FontSetter[a] = ##org.apache.poi.ss.usermodel.Font ~> Result[a, System.Error]

}

namespace SheetIO/FontSetter {

    use SheetIO.FontSetter;
    use SheetIO.FontSetter.FontSetter;

    pub def return(x: a): FontSetter[a] =
        FontSetter(_ -> Ok(x) as & Impure)

    def apply1(x: FontSetter[a], r: ##org.apache.poi.ss.usermodel.Font): Result[a, System.Error] & Impure =
        let FontSetter(f) = x;
        f(r)

    pub def mapError(f: System.Error -> System.Error, ma: FontSetter[a]): FontSetter[a] =
        FontSetter(r -> match apply1(ma, r) { 
            case Ok(p1) => Ok(p1)
            case Err(e) => Err(f(e))
        })


    pub def map(f: a -> b & ef, x: FontSetter[a]): FontSetter[b] =
        FontSetter(r ->
            match apply1(x, r) {
                case Ok(a) => {
                    let ans = f(a) as & Impure;
                    Ok(ans)
                }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: FontSetter[a -> b & ef], ma: FontSetter[a]): FontSetter[b] =
        FontSetter(r ->
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

    pub def bind(x: FontSetter[a], k: a -> FontSetter[b]): FontSetter[b] =
        FontSetter(r ->
            match apply1(x, r) {
                case Ok(a) => { let ans = k(a); apply1(ans, r) }
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(k: a -> FontSetter[b], x: FontSetter[a]): FontSetter[b] = bind(x, k)

    def liftPrim(update: ##org.apache.poi.ss.usermodel.Font ~> Result[a, System.Error]): FontSetter[a] =
        FontSetter(r -> update(r))

    pub def liftAction(fn: Unit ~> a): FontSetter[a] =
        FontSetter(_ -> let x = fn(); Ok(x))

    pub def liftActionResult(fn: Unit ~> Result[a, System.Error]): FontSetter[a] =
        FontSetter(_ -> fn())

    ///
    /// Applies the binary function `f` to the values in `o1` and `o2`.
    ///
    /// Returns an error if either `o1` or `o2` are errors.
    ///
    pub def lift2(f: (t1, t2) -> u & ef, 
                  o1: FontSetter[t1], o2: FontSetter[t2]): FontSetter[u] =
        ap(map(f, o1), o2)

    ///
    /// Applies the ternary function `f` to the values in `o1`, `o2` and `o3`.
    ///
    /// Returns an error if any of `o1`, `o2` and `o3` are errors.
    ///
    pub def lift3(f: (t1, t2, t3) -> u & ef,
                  o1: FontSetter[t1],
                  o2: FontSetter[t2],
                  o3: FontSetter[t3]): FontSetter[u] =
        ap(lift2(f, o1, o2), o3)

    ///
    /// Applies the 4-ary function `f` to the values in `o1`, `o2`, `o3` and `o4`.
    ///
    /// Returns an error if any of `o1`, `o2`, `o3` and `o4` are errors.
    ///
    pub def lift4(f: (t1, t2, t3, t4) -> u & ef,
                  o1: FontSetter[t1],
                  o2: FontSetter[t2],
                  o3: FontSetter[t3],
                  o4: FontSetter[t4]): FontSetter[u] =
        ap(lift3(f, o1, o2, o3), o4)

    ///
    /// Applies the 5-ary function `f` to the values in `o1`, `o2`, ... `o5`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o5` are errors.
    ///
    pub def lift5(f: (t1, t2, t3, t4, t5) -> u & ef,
                  o1: FontSetter[t1],
                  o2: FontSetter[t2],
                  o3: FontSetter[t3],
                  o4: FontSetter[t4],
                  o5: FontSetter[t5]): FontSetter[u] =
        ap(lift4(f, o1, o2, o3, o4), o5)

    ///
    /// Applies the 6-ary function `f` to the values in `o1`, `o2`, ... `o6`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o6` are errors.
    ///
    pub def lift6(f: (t1, t2, t3, t4, t5, t6) -> u & ef,
                  o1: FontSetter[t1],
                  o2: FontSetter[t2],
                  o3: FontSetter[t3],
                  o4: FontSetter[t4],
                  o5: FontSetter[t5],
                  o6: FontSetter[t6]): FontSetter[u] =
        ap(lift5(f, o1, o2, o3, o4, o5), o6)

    ///
    /// Applies the 7-ary function `f` to the values in `o1`, `o2`, ... `o7`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o7` are errors.
    ///
    pub def lift7(f: (t1, t2, t3, t4, t5, t6, t7) -> u & ef,
                  o1: FontSetter[t1],
                  o2: FontSetter[t2],
                  o3: FontSetter[t3],
                  o4: FontSetter[t4],
                  o5: FontSetter[t5],
                  o6: FontSetter[t6],
                  o7: FontSetter[t7]): FontSetter[u] =
        ap(lift6(f, o1, o2, o3, o4, o5, o6), o7)

    ///
    /// Applies the 8-ary function `f` to the values in `o1`, `o2`, ... `o8`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o8` are errors.
    ///
    pub def lift8(f: (t1, t2, t3, t4, t5, t6, t7, t8) -> u & ef,
                      o1: FontSetter[t1],
                      o2: FontSetter[t2],
                      o3: FontSetter[t3],
                      o4: FontSetter[t4],
                      o5: FontSetter[t5],
                      o6: FontSetter[t6],
                      o7: FontSetter[t7],
                      o8: FontSetter[t8]): FontSetter[u] =
        ap(lift7(f, o1, o2, o3, o4, o5, o6, o7), o8)

    ///
    /// Applies the 9-ary function `f` to the values in `o1`, `o2`, ... `o9`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o9` are errors.
    ///
    pub def lift9(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9) -> u & ef,
                  o1: FontSetter[t1],
                  o2: FontSetter[t2],
                  o3: FontSetter[t3],
                  o4: FontSetter[t4],
                  o5: FontSetter[t5],
                  o6: FontSetter[t6],
                  o7: FontSetter[t7],
                  o8: FontSetter[t8],
                  o9: FontSetter[t9]): FontSetter[u] =
        ap(lift8(f, o1, o2, o3, o4, o5, o6, o7, o8), o9)

    ///
    /// Applies the 10-ary function `f` to the values in `o1`, `o2`, ... `o10`.
    ///
    /// Returns an error if any of `o1`, `o2`, ... `o10` are errors.
    ///
    pub def lift10(f: (t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) -> u & ef,
                   o1: FontSetter[t1],
                   o2: FontSetter[t2],
                   o3: FontSetter[t3],
                   o4: FontSetter[t4],
                   o5: FontSetter[t5],
                   o6: FontSetter[t6],
                   o7: FontSetter[t7],
                   o8: FontSetter[t8],
                   o9: FontSetter[t9],
                   o10: FontSetter[t10]): FontSetter[u] =
        ap(lift9(f, o1, o2, o3, o4, o5, o6, o7, o8, o9), o10)


    pub def setBold!(x: Bool): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setBold(r, x))

    pub def setItalic!(x: Bool): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setItalic(r, x))

    pub def setStrikeout!(x: Bool): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setStrikeout(r, x))
    
    pub def setColor!(x: Int16): FontSetter[Unit] = 
        liftPrim(r -> SheetIO/Internal/Font.setColor(r, x))


}