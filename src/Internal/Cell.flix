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

use System.Error.{Generic};

namespace SheetIO/Internal/Cell {


    pub def getCellType(cell: ##org.apache.poi.ss.usermodel.Cell): Result[##org.apache.poi.ss.usermodel.CellType, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.getCellType(); 
        try {
            let ans = cell.getCellType();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getCellType")) 
        }

    pub def setCellType(cell: ##org.apache.poi.ss.usermodel.Cell, 
                        typ: ##org.apache.poi.ss.usermodel.CellType): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellType(##org.apache.poi.ss.usermodel.CellType); 
        try {
            let ans = cell.setCellType(typ);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setCellType")) 
        }

     
    pub def getBooleanCellValue(cell: ##org.apache.poi.ss.usermodel.Cell): Result[Bool, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.getBooleanCellValue(); 
        try {
            let ans = cell.getBooleanCellValue();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getBooleanCellValue")) 
        }


    pub def getStringCellValue(cell: ##org.apache.poi.ss.usermodel.Cell): Result[String, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.getStringCellValue(); 
        try {
            let ans = cell.getStringCellValue();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getStringCellValue")) 
        }

    pub def getNumericCellValue(cell: ##org.apache.poi.ss.usermodel.Cell): Result[Float64, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.getNumericCellValue(); 
        try {
            let ans = cell.getNumericCellValue();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getNumericCellValue")) 
        }


    pub def getLocalDateTimeCellValue(cell: ##org.apache.poi.ss.usermodel.Cell): Result[##java.time.LocalDateTime, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.getLocalDateTimeCellValue(); 
        try {
            let ans = cell.getLocalDateTimeCellValue();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getLocalDateTimeCellValue")) 
        }


    pub def setBooleanCellValue(cell: ##org.apache.poi.ss.usermodel.Cell, value: Bool): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Bool); 
        try {
            let ans = cell.setCellValue(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setBooleanCellValue")) 
        }

    pub def setNumericCellValue(cell: ##org.apache.poi.ss.usermodel.Cell, value: Float64): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(Float64); 
        try {
            let ans = cell.setCellValue(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setNumericCellValue")) 
        }

    pub def setStringCellValue(cell: ##org.apache.poi.ss.usermodel.Cell, 
                               value: String): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(String); 
        try {
            let ans = cell.setCellValue(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setStringCellValue")) 
        }


    pub def setLocalDateTimeCellValue(cell: ##org.apache.poi.ss.usermodel.Cell, 
                                      value: ##java.time.LocalDateTime): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDateTime); 
        try {
            let ans = cell.setCellValue(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setLocalDateTimeCellValue")) 
        }

    pub def setLocalDateCellValue(cell: ##org.apache.poi.ss.usermodel.Cell, 
                                  value: ##java.time.LocalDate): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellValue(##java.time.LocalDate); 
        try {
            let ans = cell.setCellValue(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setLocalDateCellValue")) 
        }


    pub def setCellStyle(cell: ##org.apache.poi.ss.usermodel.Cell, 
                         style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.setCellStyle(##org.apache.poi.ss.usermodel.CellStyle); 
        try {
            let ans = cell.setCellStyle(style);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setCellStyle")) 
        }
}