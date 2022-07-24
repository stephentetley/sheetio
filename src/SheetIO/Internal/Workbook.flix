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


namespace SheetIO/Internal/Workbook {

    use SheetIO/Internal/PrimError.{PrimError};

    pub def createSheet(r: Region[r],
                        name: String,
                        wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.Sheet, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createSheet(String): ##org.apache.poi.ss.usermodel.Sheet & r; 
        try {
            let ans = createSheet(wb, name);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createSheet", ex)
        }


    pub def getNumberOfSheets(r: Region[r], 
                                wb: ##org.apache.poi.ss.usermodel.Workbook): Result[Int32, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getNumberOfSheets(): Int32 & r; 
        try {
            let ans = getNumberOfSheets(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getNumberOfSheets", ex)
        }

    pub def getSheetName(r: Region[r], 
                            ix: Int32,
                            wb: ##org.apache.poi.ss.usermodel.Workbook): Result[String, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetName(Int32): String & r; 
        try {
            let ans = getSheetName(wb, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheetName", ex)
        }

    pub def getSheet(r: Region[r], 
                    name: String,
                    wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.Sheet, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String): ##org.apache.poi.ss.usermodel.Sheet & r; 
        try {
            let ans = getSheet(wb, name);
            if (Object.isNull(ans))
                Err({source = "getSheet", exnMessage = "Answer is null"})
            else
                Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheet", ex)
        }

    pub def getSheetAt(r: Region[r],
                        ix: Int32,
                        wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.Sheet, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32): ##org.apache.poi.ss.usermodel.Sheet & r; 
        try {
            let ans = getSheetAt(wb, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheetAt", ex)
        }

    pub def getFontAt(r: Region[r], 
                        ix: Int32,
                        wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.Font, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getFontAt(Int32): ##org.apache.poi.ss.usermodel.Font & r; 
        try {
            let ans = getFontAt(wb, ix);
            if (Object.isNull(ans))
                Err({source = "getFontAt", exnMessage = "Answer is null"})
            else
                Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getFontAt", ex)
        }


    pub def createCellStyle(r: Region[r],
                            wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.CellStyle, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle & r; 
        try {
            let ans = createCellStyle(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createCellStyle", ex)
        }

    pub def createFont(r: Region[r],
                        wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.Font, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createFont(): ##org.apache.poi.ss.usermodel.Font & r; 
        try {
            let ans = createFont(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createFont", ex)
        }

    pub def createDataFormat(r: Region[r],
                            wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.DataFormat, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createDataFormat(): ##org.apache.poi.ss.usermodel.DataFormat & r; 
        try {
            let ans = createDataFormat(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createDataFormat", ex)
        }

}
