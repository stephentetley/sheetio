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


mod Internal.Apache.Poi.SS.Usermodel {

    /// "Immutable with Setter"

    pub enum DataFormatter(##org.apache.poi.ss.usermodel.DataFormatter)

    pub enum MutDataFormatter[_: Region](##org.apache.poi.ss.usermodel.DataFormatter)


    /// DataFormatter "builders" in POI return void ...
    pub type alias DataFormatterInit[r: Region] = MutDataFormatter[r] -> Result[String, Unit] \ r

}


mod Internal.Apache.Poi.SS.Usermodel.DataFormatter {

    use Locale.{Locale}

    use Basis.Text.DecimalFormat
    use Basis.Text.DecimalFormat.{DecimalFormat}

    use Internal.Apache.Poi.SS.Usermodel.DataFormatter
    use Internal.Apache.Poi.SS.Usermodel.DataFormatter.{DataFormatter}
    use Internal.Apache.Poi.SS.Usermodel.MutDataFormatter
    use Internal.Apache.Poi.SS.Usermodel.MutDataFormatter.{MutDataFormatter}
    use Internal.Apache.Poi.SS.Usermodel.DataFormatterInit
    use Internal.Apache.Poi.SS.Usermodel.Cell
    use Internal.Apache.Poi.SS.Usermodel.Cell.{Cell}
    


    pub def getHH(df: DataFormatter): ##org.apache.poi.ss.usermodel.DataFormatter =
        let DataFormatter(df1) = df;
        df1

    pub def new(_: Region[r], init: DataFormatterInit[r]): Result[String, DataFormatter] \ r = 
        import new org.apache.poi.ss.usermodel.DataFormatter(): ##org.apache.poi.ss.usermodel.DataFormatter \ {} as new1;
        let df1 = new1();
        MutDataFormatter(df1) |> init |> Result.map(_ -> DataFormatter(df1))

    
    ///
    /// Returns a new mutable DataFormatter.
    ///
    pub def newWithLocale(_: Region[r], loc: Locale, init: DataFormatterInit[r]): Result[String, DataFormatter] \ r = 
        import new org.apache.poi.ss.usermodel.DataFormatter(##java.util.Locale): ##org.apache.poi.ss.usermodel.DataFormatter \ r as new1;
        let Locale(loc1) = loc;
        let df1 = new1(loc1);
        MutDataFormatter(df1) |> init |> Result.map(_ -> DataFormatter(df1))


    /// This is a "service" ...
    pub def formatCellValue(cell: Cell, df: DataFormatter): String = 
        import org.apache.poi.ss.usermodel.DataFormatter.formatCellValue(##org.apache.poi.ss.usermodel.Cell): String \ {};
        let DataFormatter(df1) = df;
        let Cell(cell1) = cell;
        formatCellValue(df1, cell1)

    /// Setters

        
    pub def setDefaultNumberFormat!(fmt: ##java.text.Format, df: MutDataFormatter[r]): Unit \ r =
        import org.apache.poi.ss.usermodel.DataFormatter.setDefaultNumberFormat(##java.text.Format): Unit \ r;
        let MutDataFormatter(df1) = df;
        setDefaultNumberFormat(df1, fmt)

    /// TO REMOVE... once `platform-basis` > 0.2.0
    pub type alias DecimalFormatBuilder = ##java.text.DecimalFormat -> Unit \ IO

    /// Operates on a `DecimalFormat` ...
    pub def setExcelStyleRoundingMode!(): DecimalFormatBuilder = df ->
        import static org.apache.poi.ss.usermodel.DataFormatter.setExcelStyleRoundingMode(##java.text.DecimalFormat): Unit \ IO;
        setExcelStyleRoundingMode(df)

}

