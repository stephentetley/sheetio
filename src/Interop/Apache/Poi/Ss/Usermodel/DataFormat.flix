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

    /// Mutable* / Interface

    pub enum DataFormat[_: Region](##org.apache.poi.ss.usermodel.DataFormat)

    instance Scoped[DataFormat] {
        pub def regionOf(_: DataFormat[r]): Region[r] = () as Region[r]
    }

}


namespace Interop/Apache/Poi/Ss/Usermodel/DataFormat {

    use Interop/Apache/Poi/Ss/Usermodel.DataFormat;
    use Interop/Apache/Poi/Ss/Usermodel.DataFormat.{DataFormat};

    pub def getHH(df: DataFormat[r]): ##org.apache.poi.ss.usermodel.DataFormat =
        let DataFormat(df1) = df;
        df1


    pub def getFormatAt(ix: Int16, df: DataFormat[r]): Option[String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.DataFormat.getFormat(Int16): String \ r; 
        let DataFormat(df1) = df;
        getFormat(df1, ix) |> Object.toOption

    /// Creates a new format if not found and returns its index.
    pub def getFormatNamed!(format: String, df: DataFormat[r]): Int16 \ Read(r) = 
        import org.apache.poi.ss.usermodel.DataFormat.getFormat(String): Int16 \ r; 
        let DataFormat(df1) = df;
        getFormat(df1, format)

}

