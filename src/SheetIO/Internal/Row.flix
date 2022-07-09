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


    use SheetIO/Internal/PrimError.{PrimError};
    
    pub def getRowNum(_: Region[r], row: ##org.apache.poi.ss.usermodel.Row): Int32 \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getRowNum(): Int32 & r; 
        getRowNum(row)
    
    pub def getCell(r: Region[r], 
                    ix: Int32, 
                    row: ##org.apache.poi.ss.usermodel.Row): Result[##org.apache.poi.ss.usermodel.Cell, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getCell(Int32): ##org.apache.poi.ss.usermodel.Cell & r; 
        try {
            let ans = getCell(row, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getCell", ex) 
        }

    pub def getSheet(r: Region[r], 
                    row: ##org.apache.poi.ss.usermodel.Row): Result[##org.apache.poi.ss.usermodel.Sheet, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getSheet(): ##org.apache.poi.ss.usermodel.Sheet & r; 
        try {
            let ans = getSheet(row);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheet", ex)
        }

    pub def createCell!(r: Region[r], 
                        ix: Int32,
                        row: ##org.apache.poi.ss.usermodel.Row): Result[##org.apache.poi.ss.usermodel.Cell, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Row.createCell(Int32): ##org.apache.poi.ss.usermodel.Cell & r; 
        try {
            let ans = createCell(row, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createCell!", ex)
        }


    pub def getFirstCellNum(r: Region[r], 
                            row: ##org.apache.poi.ss.usermodel.Row): Result[Int32, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getFirstCellNum(): Int16 & r; 
        try {
            let ans = getFirstCellNum(row);
            Ok(Int16.toInt32(ans))
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getFirstCellNum", ex)
        }


    /// As per Java - index of last cell **plus 1**.
    pub def getLastCellNum(r: Region[r], 
                            row: ##org.apache.poi.ss.usermodel.Row): Result[Int32, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getLastCellNum(): Int16 & r; 
        try {
            let ans = getLastCellNum(row);
            Ok(Int16.toInt32(ans))
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getLastCellNum", ex)
        }


    /// Number of cells used - do not use for trying to find end.
    pub def getPhysicalNumberOfCells(r: Region[r], 
                                    row: ##org.apache.poi.ss.usermodel.Row): Result[Int32, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getPhysicalNumberOfCells(): Int32 & r; 
        try {
            let ans = getPhysicalNumberOfCells(row);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getPhysicalNumberOfCells", ex)
        }

}
