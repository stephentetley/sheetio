/*
 * Copyright 2022 Stephen Tetley
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


namespace Interop/Apache/Poi/Ss/Usermodel {

    /// Mutable / Class

    pub enum DataFormatter[_: Region](##org.apache.poi.ss.usermodel.DataFormatter)

    instance Scoped[DataFormatter] {
        pub def regionOf(_: DataFormatter[r]): Region[r] = () as Region[r]
    }
}


namespace Interop/Apache/Poi/Ss/Usermodel/DataFormatter {

    use Interop/Apache/Poi/Ss/Usermodel.DataFormatter;
    use Interop/Apache/Poi/Ss/Usermodel.DataFormatter.{DataFormatter};
    use Interop/Apache/Poi/Ss/Usermodel.Cell;
    use Interop/Apache/Poi/Ss/Usermodel.Cell.{Cell};


    pub def getHH(df: DataFormatter[r]): ##org.apache.poi.ss.usermodel.DataFormatter =
        let DataFormatter(df1) = df;
        df1

    pub def new(_: Region[r]): DataFormatter[r] \ Write(r) = 
        import new org.apache.poi.ss.usermodel.DataFormatter(): ##org.apache.poi.ss.usermodel.DataFormatter \ r as newDataFormatter;
        DataFormatter(newDataFormatter())


    // ///
    // /// Returns a new mutable DataFormatter.
    // ///
    // pub def newWithLocale(_: Region[r], x: Locale): ##org.apache.poi.ss.usermodel.DataFormatter \ Write(r) = 
    //     import new org.apache.poi.ss.usermodel.DataFormatter(##java.util.Locale): ##org.apache.poi.ss.usermodel.DataFormatter \ r as newDataFormatter;
    //     let Locale(x1) = x;
    //     newDataFormatter(Interop/Util/Locale.getHH(x1))


    /// TODO - it would be useful to allow users to make custom DataFormatters...

    /// This is a "getter" despite its name.
    pub def formatCellValue(cell: Cell[r], df: DataFormatter[r]): String \ Read(r) = 
        import org.apache.poi.ss.usermodel.DataFormatter.formatCellValue(##org.apache.poi.ss.usermodel.Cell): String \ r; 
        let DataFormatter(df1) = df;
        let Cell(cell1) = cell;
        formatCellValue(df1, cell1)


}

