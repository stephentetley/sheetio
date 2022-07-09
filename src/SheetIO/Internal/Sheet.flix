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

namespace SheetIO/Internal/Sheet {

    use SheetIO/Internal/PrimError.{PrimError};

    pub def getSheetName(r: Region[r], sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[String, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Sheet.getSheetName(): String & r; 
        try {
            let ans = getSheetName(sheet);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheetName", ex) 
        }

    pub def getFirstRowNum(r: Region[r], 
                            sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Sheet.getFirstRowNum(): Int32 & r; 
        try {
            let ans = getFirstRowNum(sheet);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getFirstRowNum", ex)
        }

    /// Gets the last row number - this row could be "empty" if cells in the row have been written to but then cleared.
    pub def getLastRowNum(r: Region[r], 
                            sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, PrimError] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getLastRowNum(): Int32 & Impure;
        try {
            let ans = getLastRowNum(sheet);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getLastRowNum", ex)
        }

    /// Number of rows used - do not use for trying to find end.
    pub def getPhysicalNumberOfRows(r: Region[r], 
                                    sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, PrimError] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getPhysicalNumberOfRows(): Int32 & Impure; 
        try {
            let ans = getPhysicalNumberOfRows(sheet);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getPhysicalNumberOfRows", ex)
        }

    pub def getWorkbook(r: Region[r], 
                        sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[##org.apache.poi.ss.usermodel.Workbook, PrimError] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getWorkbook(): ##org.apache.poi.ss.usermodel.Workbook & Impure;
        try {
            let ans = getWorkbook(sheet);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getWorkbook", ex)
        }


    pub def getRow(r: Region[r],
                    ix: Int32,
                    sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[##org.apache.poi.ss.usermodel.Row, PrimError] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getRow(Int32): ##org.apache.poi.ss.usermodel.Row & Impure; 
        try {
            let ans = getRow(sheet, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getRow", ex)
        }

    pub def createRow(r: Region[r],
                        ix: Int32,
                        sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[##org.apache.poi.ss.usermodel.Row, PrimError] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.createRow(Int32): ##org.apache.poi.ss.usermodel.Row & Impure; 
        try {
            let ans = createRow(sheet, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createRow", ex)
        }

}
