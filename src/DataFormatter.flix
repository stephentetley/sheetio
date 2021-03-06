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
    /// CellType
    ///
    pub opaque type DataFormatter = ##org.apache.poi.ss.usermodel.DataFormatter

}

namespace SheetIO/DataFormatter {

    use SheetIO.DataFormatter;
    use SheetIO.DataFormatter.{DataFormatter};

    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def new(): DataFormatter & Impure =
        import new org.apache.poi.ss.usermodel.DataFormatter() as newDataFormatter;
        DataFormatter(newDataFormatter())

    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def newWithLocale(x: I18n.Locale): DataFormatter & Impure =
        use I18n.Locale.{Locale};
        import new org.apache.poi.ss.usermodel.DataFormatter(##java.util.Locale) as newDataFormatter;
        let Locale(x1) = x;
        DataFormatter(newDataFormatter(x1))
        
}