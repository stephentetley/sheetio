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

namespace SheetIO/Internal/Row {

    pub def getRowNum(row: ##org.apache.poi.ss.usermodel.Row): Int32 & Impure = 
        import org.apache.poi.ss.usermodel.Row.getRowNum(): Int32 & Impure; 
        getRowNum(row)
    
    pub def getCell(row: ##org.apache.poi.ss.usermodel.Row, 
                    ix: Int32): Result[##org.apache.poi.ss.usermodel.Cell, String] & Impure = 
        import org.apache.poi.ss.usermodel.Row.getCell(Int32): ##org.apache.poi.ss.usermodel.Cell & Impure; 
        try {
            let ans = getCell(row, ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getCell")
        }

    pub def getSheet(row: ##org.apache.poi.ss.usermodel.Row): Result[##org.apache.poi.ss.usermodel.Sheet, String] & Impure = 
        import org.apache.poi.ss.usermodel.Row.getSheet(): ##org.apache.poi.ss.usermodel.Sheet & Impure; 
        try {
            let ans = getSheet(row);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getSheet")
        }

    pub def createCell(row: ##org.apache.poi.ss.usermodel.Row, 
                       ix: Int32): Result[##org.apache.poi.ss.usermodel.Cell, String] & Impure = 
        import org.apache.poi.ss.usermodel.Row.createCell(Int32): ##org.apache.poi.ss.usermodel.Cell & Impure; 
        try {
            let ans = createCell(row, ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("createCell")
        }


    pub def getFirstCellNum(row: ##org.apache.poi.ss.usermodel.Row): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Row.getFirstCellNum(): Int16 & Impure; 
        try {
            let ans = getFirstCellNum(row);
            Ok(Int16.toInt32(ans))
        } catch {
            case _: ##java.lang.Exception => Err("getFirstCellNum")
        }


    /// As per Java - index of last cell **plus 1**.
    pub def getLastCellNum(row: ##org.apache.poi.ss.usermodel.Row): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Row.getLastCellNum(): Int16 & Impure; 
        try {
            let ans = getLastCellNum(row);
            Ok(Int16.toInt32(ans))
        } catch {
            case _: ##java.lang.Exception => Err("getLastCellNum")
        }


    /// Number of cells used - do not use for trying to find end.
    pub def getPhysicalNumberOfCells(row: ##org.apache.poi.ss.usermodel.Row): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Row.getPhysicalNumberOfCells(): Int32 & Impure; 
        try {
            let ans = getPhysicalNumberOfCells(row);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getPhysicalNumberOfCells")
        }

}
