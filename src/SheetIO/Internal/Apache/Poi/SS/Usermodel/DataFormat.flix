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


mod SheetIO.Internal.Apache.Poi.SS.Usermodel {

    /// Mutable*  / Interface
    ///
    /// * Interface is immutable but subclass XSSFDataFormat is mutable.
    pub enum DataFormat[_: Region](##org.apache.poi.ss.usermodel.DataFormat)


}


mod SheetIO.Internal.Apache.Poi.SS.Usermodel.DataFormat {

    use SheetIO.Internal.Apache.Poi.SS.Usermodel.DataFormat
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.DataFormat.{DataFormat}



    pub def getFormatAt(ix: Int16, df: DataFormat[r]): Option[String] \ r = 
        import org.apache.poi.ss.usermodel.DataFormat.getFormat(Int16): String \ r; 
        let DataFormat(df1) = df;
        getFormat(df1, ix) |> Object.toOption

    /// Creates a new format if not found and returns its index.
    pub def getFormatNamed!(format: String, df: DataFormat[r]): Int16 \ r = 
        import org.apache.poi.ss.usermodel.DataFormat.getFormat(String): Int16 \ r; 
        let DataFormat(df1) = df;
        getFormat(df1, format)

}

