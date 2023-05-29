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

mod Interop.Apache.Poi.SS.Usermodel {

    /// Mutable / Interface

    pub enum MutCell[_: Region](##org.apache.poi.ss.usermodel.Cell)


}

mod Interop.Apache.Poi.SS.Usermodel.MutCell {

    use Interop.Apache.Poi.SS.Usermodel.MutCell
    use Interop.Apache.Poi.SS.Usermodel.MutCell.{MutCell}
    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Interop.Apache.Poi.SS.Usermodel.CellType
    use Interop.Apache.Poi.SS.Usermodel.CellType.{CellType}

    pub def getHH(cell: MutCell[r]): ##org.apache.poi.ss.usermodel.Cell =
        let MutCell(cell1) = cell;
        cell1


    pub def wrapHH(_: Region[r], prim:  ##org.apache.poi.ss.usermodel.Cell): MutCell[r] =
        MutCell(prim)



    pub def setCellValueBoolean!(value: Bool, cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Bool): Unit \ r; 
        let MutCell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueDouble!(value: Float64,
                                cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Float64): Unit \ r; 
        let MutCell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueString!(value: String,
                                cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(String): Unit \ r;
        let MutCell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueLocalDate!(value: ##java.time.LocalDate, cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDate): Unit \ r;
        let MutCell(cell1) = cell;
        setCellValue(cell1, value)


    pub def setCellValueLocalDateTime!(value: ##java.time.LocalDateTime, cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDateTime): Unit \ r;
        let MutCell(cell1) = cell;
        setCellValue(cell1, value)



    pub def setBlank!(cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setBlank(): Unit \ r;
        let MutCell(cell1) = cell;
        setBlank(cell1)


    pub def setCellStyle!(style: CellStyle[r], cell: MutCell[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.Cell.setCellStyle(##org.apache.poi.ss.usermodel.CellStyle): Unit \ r;
        let MutCell(cell1) = cell;
        let CellStyle(style1) = style;
        setCellStyle(cell1, style1)

   pub def getCellStyle(cell: MutCell[r]): CellStyle[r] \ r = 
        import org.apache.poi.ss.usermodel.Cell.getCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ r; 
        let MutCell(cell1) = cell;
        getCellStyle(cell1) |> CellStyle


}

