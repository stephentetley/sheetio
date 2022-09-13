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

    /// Mutable / Interface

    pub enum Cell[_: Region](##org.apache.poi.ss.usermodel.Cell)

    instance Scoped[Cell] {
        pub def regionOf(_: Cell[r]): Region[r] = () as Region[r]
    }
}

namespace Interop/Apache/Poi/Ss/Usermodel/Cell {

    use Interop/Apache/Poi/Ss/Usermodel.Cell;
    use Interop/Apache/Poi/Ss/Usermodel.Cell.{Cell};

    pub def getHH(cell: Cell[r]): ##org.apache.poi.ss.usermodel.Cell =
        let Cell(cell1) = cell;
        cell1

    pub def getCellType(cell: Cell[r]): ##org.apache.poi.ss.usermodel.CellType \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getCellType(): ##org.apache.poi.ss.usermodel.CellType \ r; 
        let Cell(cell1) = cell;
        getCellType(cell1)


    pub def getBooleanCellValue(cell: Cell[r]): Result[Bool, String] \ Read(r) =  
        try {
            import org.apache.poi.ss.usermodel.Cell.getBooleanCellValue(): Bool \ r;
            let Cell(cell1) = cell;
            getBooleanCellValue(cell1) |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


    pub def getStringCellValue(cell: Cell[r]): Result[String, String] \ Read(r) = 
        try {
            import org.apache.poi.ss.usermodel.Cell.getStringCellValue(): String \ r; 
            let Cell(cell1) = cell;
            let ans = getStringCellValue(cell1);
            if (not Object.isNull(ans)) Ok(ans) else Ok("")
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }
/*
    pub def getNumericCellValue(r: Region[r],
                                cell: Cell[r]): Result[Float64, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getNumericCellValue(): Float64 \ r; 
        try {
            let ans = getNumericCellValue(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getNumericCellValue", ex)
        }


    pub def getLocalDateTimeCellValue(r: Region[r],
                                        cell: Cell[r]): Result[##java.time.LocalDateTime, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getLocalDateTimeCellValue(): ##java.time.LocalDateTime \ r; 
        try {
            let ans = getLocalDateTimeCellValue(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getLocalDateTimeCellValue", ex)
        }


    pub def setCellValueBoolean!(r: Region[r],
                                value: Bool,
                                cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Bool): Unit \ r; 
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setCellValueBoolean!", ex)
        }

    pub def setCellValueDouble!(r: Region[r], 
                                value: Float64,
                                cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Float64): Unit \ r; 
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setCellValueDouble!", ex)
        }

    pub def setCellValueString!(r: Region[r], 
                                value: String,
                                cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(String): Unit \ r;
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setCellValueString!", ex)
        }


    pub def setCellValueLocalDate!(r: Region[r], 
                                    value: ##java.time.LocalDate,
                                    cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDate): Unit \ r;
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setCellValueLocalDate!", ex)
        }

    pub def setCellValueLocalDateTime!(r: Region[r], 
                                        value: ##java.time.LocalDateTime,
                                        cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDateTime): Unit \ r;
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setCellValueLocalDateTime!", ex)
        }



    pub def setBlank!(r: Region[r], 
                        cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setBlank(): Unit \ r;
        try {
            let ans = setBlank(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setBlank!", ex)
        }


    pub def setCellStyle!(r: Region[r], 
                            style: ##org.apache.poi.ss.usermodel.CellStyle,
                            cell: Cell[r]): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellStyle(##org.apache.poi.ss.usermodel.CellStyle): Unit \ r; 
        try {
            let ans = setCellStyle(cell, style);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setCellStyle!", ex)
        }

   pub def getCellStyle(r: Region[r],
                        cell: Cell[r]): Result[##org.apache.poi.ss.usermodel.CellStyle, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ r; 
        try {
            let ans = getCellStyle(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getCellStyle", ex)
        }
*/
}

