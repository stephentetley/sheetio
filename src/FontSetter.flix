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

    /// def apply1(x: FontSetter[a], r: ##org.apache.poi.ss.usermodel.Font): Result[a, System.Error] & Impure =
    ///     let FontSetter(f) = x;
    ///     f(r)

    pub def setBold!(x: Bool): FontSetter[Unit] = 
        FontSetter(r -> SheetIO/Internal/Font.setBold(r, x))
        
}