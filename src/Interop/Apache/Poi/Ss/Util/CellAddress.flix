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

namespace Interop/Apache/Poi/Ss/Util {

    /// Immutable

    pub enum CellAddress(##org.apache.poi.ss.util.CellAddress)


}

namespace Interop/Apache/Poi/Ss/Util/CellAddress {

    use Interop/Apache/Poi/Ss/Util.CellAddress;
    use Interop/Apache/Poi/Ss/Util.CellAddress.{CellAddress};

    pub def cellAddress(row: {row = Int32}, column: {column = Int32}): CellAddress = 
        import new org.apache.poi.ss.util.CellAddress(Int32, Int32): ##org.apache.poi.ss.util.CellAddress \ {} as newCellAddress; 
        newCellAddress(row.row, column.column) |> CellAddress

    pub def getRow(ca: CellAddress): Int32 = 
        import org.apache.poi.ss.util.CellAddress.getRow(): Int32 \ {}; 
        let CellAddress(ca1) = ca;
        getRow(ca1)

    pub def getColumn(ca: CellAddress): Int32 = 
        import org.apache.poi.ss.util.CellAddress.getColumn(): Int32 \ {}; 
        let CellAddress(ca1) = ca;
        getColumn(ca1)

}

