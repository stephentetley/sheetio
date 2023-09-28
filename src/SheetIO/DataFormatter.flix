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

    ///
    /// DataFormatter
    ///
    ///
    pub type alias DataFormatter = Internal.Apache.Poi.SS.Usermodel.DataFormatter

}

mod SheetIO.DataFormatter {


    use SheetIO.DataFormatter
    use Internal.Apache.Poi.SS.Usermodel.DataFormatter.{DataFormatter}

/*
    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def new(rc: Region[r]): DataFormatter[r] \ r =
        Internal.Apache.Poi.SS.Usermodel.DataFormatter.new(rc)

    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def newWithLocale(rc: Region[r], loc: Locale): DataFormatter[r] \ r =
        Internal.Apache.Poi.SS.Usermodel.DataFormatter.newWithLocale(rc, loc)
*/      
}
