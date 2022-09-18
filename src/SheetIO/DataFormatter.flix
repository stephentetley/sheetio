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

    ///
    /// DataFormatter
    ///
    pub type alias DataFormatter[r: Region] = Interop/Apache/Poi/SS/Usermodel.DataFormatter[r]

}

namespace SheetIO/DataFormatter {


    use Basis/I18n.Locale;

    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};

    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def new(r: Region[r]): DataFormatter[r] \ Write(r) =
        Interop/Apache/Poi/SS/Usermodel/DataFormatter.new(r)

    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def newWithLocale(r: Region[r], loc: Locale): DataFormatter[r] \ Write(r) =
        let loc1 = Basis/I18n/Locale.getLocale(loc);
        Interop/Apache/Poi/SS/Usermodel/DataFormatter.newWithLocale(r, loc1)
        
}
