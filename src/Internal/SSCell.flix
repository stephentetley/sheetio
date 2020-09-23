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

namespace SheetReader/Internal/SSCell {


    pub def getCellType(cell: ##org.apache.poi.ss.usermodel.Cell): Result[##org.apache.poi.ss.usermodel.CellType, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Cell.getCellType(); 
        try {
            let ans = cell.getCellType();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getCellType")) 
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

}