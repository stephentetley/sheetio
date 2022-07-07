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

namespace SheetIO/Internal/Cell {


    pub def getCellType(_: Region[r], 
                        cell: ##org.apache.poi.ss.usermodel.Cell): Result[##org.apache.poi.ss.usermodel.CellType, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getCellType(): ##org.apache.poi.ss.usermodel.CellType & r; 
        try {
            let ans = getCellType(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getCellType: ${getMessage(ex)}") 
        }

    /// /// Deprecated...
    /// pub def setCellType(cell: ##org.apache.poi.ss.usermodel.Cell, 
    ///                     typ: ##org.apache.poi.ss.usermodel.CellType): Result[Unit, String] & Impure = 
    ///     import org.apache.poi.ss.usermodel.Cell.setCellType(##org.apache.poi.ss.usermodel.CellType): Unit & Impure; 
    ///     try {
    ///         let ans = setCellType(cell, typ);
    ///         Ok(ans)
    ///     } catch {
    ///         case _: ##java.lang.Exception => Err("setCellType")
    ///         case ex: ##java.lang.Exception => 
    ///             import java.lang.Throwable.getMessage(): String & r;                
    ///             Err("getCellType: ${getMessage(ex)}")
    ///     }

     
    pub def getBooleanCellValue(_: Region[r],
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[Bool, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getBooleanCellValue(): Bool & r; 
        try {
            let ans = getBooleanCellValue(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getBooleanCellValue: ${getMessage(ex)}")
        }


    pub def getStringCellValue(_: Region[r],
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[String, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getStringCellValue(): String & r; 
        try {
            let ans = getStringCellValue(cell);
            if (not Object.isNull(ans)) Ok(ans) else Ok("")
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getStringCellValue: ${getMessage(ex)}")
        }

    pub def getNumericCellValue(_: Region[r],
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[Float64, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getNumericCellValue(): Float64 & r; 
        try {
            let ans = getNumericCellValue(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getNumericCellValue: ${getMessage(ex)}")
        }


    pub def getLocalDateTimeCellValue(_: Region[r],
                                        cell: ##org.apache.poi.ss.usermodel.Cell): Result[##java.time.LocalDateTime, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Cell.getLocalDateTimeCellValue(): ##java.time.LocalDateTime & r; 
        try {
            let ans = getLocalDateTimeCellValue(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getLocalDateTimeCellValue: ${getMessage(ex)}")
        }


    pub def setCellValueBoolean!(_: Region[r],
                                value: Bool,
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Bool): Unit & r; 
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setCellValueBoolean!: ${getMessage(ex)}")
        }

    pub def setCellValueDouble!(_: Region[r], 
                                value: Float64,
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Float64): Unit & r; 
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setCellValueDouble!: ${getMessage(ex)}")
        }

    pub def setCellValueString!(_: Region[r], 
                                value: String,
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(String): Unit & r;
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setStringCellValue")
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setCellValueString!: ${getMessage(ex)}")
        }


    pub def setCellValueLocalDate!(_: Region[r], 
                                    value: ##java.time.LocalDate,
                                    cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDate): Unit & r;
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err(" setCellValueLocalDate!: ${getMessage(ex)}")
        }

    pub def setCellValueLocalDateTime!(_: Region[r], 
                                        value: ##java.time.LocalDateTime,
                                        cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDateTime): Unit & r;
        try {
            let ans = setCellValue(cell, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setCellValueLocalDateTime!: ${getMessage(ex)}")
        }



    pub def setBlank!(_: Region[r], 
                        cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setBlank(): Unit & r;
        try {
            let ans = setBlank(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err(" setBlank!: ${getMessage(ex)}")
        }


    pub def setCellStyle!(_: Region[r], 
                            style: ##org.apache.poi.ss.usermodel.CellStyle,
                            cell: ##org.apache.poi.ss.usermodel.Cell): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.setCellStyle(##org.apache.poi.ss.usermodel.CellStyle): Unit & r; 
        try {
            let ans = setCellStyle(cell, style);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setCellStyle!: ${getMessage(ex)}")
        }

   pub def getCellStyle(_: Region[r],
                        cell: ##org.apache.poi.ss.usermodel.Cell): Result[##org.apache.poi.ss.usermodel.CellStyle, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Cell.getCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle & r; 
        try {
            let ans = getCellStyle(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getCellStyle: ${getMessage(ex)}")
        }

}

