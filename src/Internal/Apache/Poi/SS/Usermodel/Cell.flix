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

    /// Mutable / Interface

    pub enum Cell(##org.apache.poi.ss.usermodel.Cell)


}

mod Internal.Apache.Poi.SS.Usermodel.Cell {

    use Internal.Apache.Poi.SS.Usermodel.Cell
    use Internal.Apache.Poi.SS.Usermodel.Cell.{Cell}
    use Internal.Apache.Poi.SS.Usermodel.CellStyle
    use Internal.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Internal.Apache.Poi.SS.Usermodel.CellType
    use Internal.Apache.Poi.SS.Usermodel.CellType.{CellType}

    pub def getHH(cell: Cell): ##org.apache.poi.ss.usermodel.Cell =
        let Cell(cell1) = cell;
        cell1


    pub def wrapHH(prim:  ##org.apache.poi.ss.usermodel.Cell): Cell =
        Cell(prim)


    pub def getCellType(cell: Cell): CellType = 
        import org.apache.poi.ss.usermodel.Cell.getCellType(): ##org.apache.poi.ss.usermodel.CellType \ {}; 
        let Cell(cell1) = cell;
        getCellType(cell1) |> CellType


    pub def getBooleanCellValue(cell: Cell): Result[String, Bool] =  
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getBooleanCellValue(): Bool \ {};
            let Cell(cell1) = cell;
            getBooleanCellValue(cell1)
        })


    pub def getStringCellValue(cell: Cell): Result[String, String] = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getStringCellValue(): String \ {}; 
            let Cell(cell1) = cell;
            let ans = getStringCellValue(cell1);
            if (not Object.isNull(ans)) ans else ""
        })

    pub def getNumericCellValue(cell: Cell): Result[String, Float64] = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getNumericCellValue(): Float64 \ {};
            let Cell(cell1) = cell;
            getNumericCellValue(cell1)
        })


    pub def getLocalDateTimeCellValue(cell: Cell): Result[String, ##java.time.LocalDateTime] = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getLocalDateTimeCellValue(): ##java.time.LocalDateTime \ {};
            let Cell(cell1) = cell;
            getLocalDateTimeCellValue(cell1)
        })

//    pub def getCellStyle(cell: Cell): CellStyle[r] = 
//         import org.apache.poi.ss.usermodel.Cell.getCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ {}; 
//         let Cell(cell1) = cell;
//         getCellStyle(cell1) |> CellStyle


}

