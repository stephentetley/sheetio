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

    /// Mutable objects with regions...
    
    pub enum Sheet[_: Region](##org.apache.poi.ss.usermodel.Sheet)

    pub enum Row[_: Region](##org.apache.poi.ss.usermodel.Row)

    pub enum CellStyle[_: Region](##org.apache.poi.ss.usermodel.CellStyle)

    pub enum DataFormat[_: Region](##org.apache.poi.ss.usermodel.DataFormat)

    pub enum Font[_: Region](##org.apache.poi.ss.usermodel.Font)




    instance CollectionInterop.Marshal[Row[r]] {
        pub def toJavaObject(x: Row[r]): ##java.lang.Object = 
            let Row(r1) = x;
            r1 as ##java.lang.Object

        pub def fromJavaObject(o: ##java.lang.Object): Row[r] = 
            Row(o as ##org.apache.poi.ss.usermodel.Row)

    }

}

namespace SheetIO/Common {

    use SheetIO.Row;
    use SheetIO.Row.{Row};

    pub def getRow(row: Row[_]): ##org.apache.poi.ss.usermodel.Row =
        let Row(row1) = row;
        row1

}
