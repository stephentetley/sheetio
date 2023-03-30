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

namespace Interop/Apache/Poi/SS/Usermodel {

    /// Mutable / Interface

    pub enum Cell[_: Region](##org.apache.poi.ss.usermodel.Cell)


}

namespace Interop/Apache/Poi/SS/Usermodel/Cell {

    use Interop/Apache/Poi/SS/Usermodel.Cell;
    use Interop/Apache/Poi/SS/Usermodel.Cell.{Cell};
    use Interop/Apache/Poi/SS/Usermodel.CellStyle;
    use Interop/Apache/Poi/SS/Usermodel.CellStyle.{CellStyle};
    use Interop/Apache/Poi/SS/Usermodel.CellType;
    use Interop/Apache/Poi/SS/Usermodel.CellType.{CellType};

    pub def getHH(cell: Cell[r]): ##org.apache.poi.ss.usermodel.Cell =
        let Cell(cell1) = cell;
        cell1


    pub def wrapHH(_: Region[r], prim:  ##org.apache.poi.ss.usermodel.Cell): Cell[r] =
        Cell(prim)


    pub def getCellType(cell: Cell[r]): CellType \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getCellType(): ##org.apache.poi.ss.usermodel.CellType \ r; 
        let Cell(cell1) = cell;
        getCellType(cell1) |> CellType


    pub def getBooleanCellValue(cell: Cell[r]): Result[String, Bool] \ Read(r) =  
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getBooleanCellValue(): Bool \ r;
            let Cell(cell1) = cell;
            getBooleanCellValue(cell1)
        })


    pub def getStringCellValue(cell: Cell[r]): Result[String, String] \ Read(r) = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getStringCellValue(): String \ r; 
            let Cell(cell1) = cell;
            let ans = getStringCellValue(cell1);
            if (not Object.isNull(ans)) ans else ""
        })

    pub def getNumericCellValue(cell: Cell[r]): Result[String, Float64] \ Read(r) = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getNumericCellValue(): Float64 \ r;
            let Cell(cell1) = cell;
            getNumericCellValue(cell1)
        })


    pub def getLocalDateTimeCellValue(cell: Cell[r]): Result[String, ##java.time.LocalDateTime] \ Read(r) = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Cell.getLocalDateTimeCellValue(): ##java.time.LocalDateTime \ r;
            let Cell(cell1) = cell;
            getLocalDateTimeCellValue(cell1)
        })


    pub def setCellValueBoolean!(value: Bool, cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Bool): Unit \ r; 
        let Cell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueDouble!(value: Float64,
                                cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Float64): Unit \ r; 
        let Cell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueString!(value: String,
                                cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(String): Unit \ r;
        let Cell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueLocalDate!(value: ##java.time.LocalDate, cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDate): Unit \ r;
        let Cell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueLocalDateTime!(value: ##java.time.LocalDateTime, cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDateTime): Unit \ r;
        let Cell(cell1) = cell;
        setCellValue(cell1, value)



    pub def setBlank!(cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setBlank(): Unit \ r;
        let Cell(cell1) = cell;
        setBlank(cell1)


    pub def setCellStyle!(style: CellStyle[r], cell: Cell[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellStyle(##org.apache.poi.ss.usermodel.CellStyle): Unit \ r;
        let Cell(cell1) = cell;
        let CellStyle(style1) = style;
        setCellStyle(cell1, style1)

   pub def getCellStyle(cell: Cell[r]): CellStyle[r] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ r; 
        let Cell(cell1) = cell;
        getCellStyle(cell1) |> CellStyle


}

