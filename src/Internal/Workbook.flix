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


    pub def createSheet(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                        name: String): Result[##org.apache.poi.ss.usermodel.Sheet, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.createSheet(String); 
        try {
            let ans = createSheet(wb, name);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("createSheet")
        }


    pub def getNumberOfSheets(wb: ##org.apache.poi.ss.usermodel.Workbook): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getNumberOfSheets(); 
        try {
            let ans = getNumberOfSheets(wb);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getNumberOfSheets")
        }

    pub def getSheetName(wb: ##org.apache.poi.ss.usermodel.Workbook, ix: Int32): Result[String, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetName(Int32); 
        try {
            let ans = getSheetName(wb, ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getSheetName")
        }

    pub def getSheet(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                     name: String): Result[##org.apache.poi.ss.usermodel.Sheet, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String); 
        try {
            let ans = getSheet(wb, name);
            if (Object.isNull(ans))
                Err("getSheet('${name}') is null")
            else
                Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getSheet")
        }

    pub def getSheetAt(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                       ix: Int32): Result[##org.apache.poi.ss.usermodel.Sheet, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32); 
        try {
            let ans = getSheetAt(wb, ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getSheetAt") 
        }

    pub def getFontAt(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                     idx: Int32): Result[##org.apache.poi.ss.usermodel.Font, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getFontAt(Int32); 
        try {
            let ans = getFontAt(wb, idx);
            if (Object.isNull(ans))
                Err("getFontAt('${idx}') answer is null")
            else
                Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getFontAt")
        }


    pub def createCellStyle(wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.CellStyle, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.createCellStyle(); 
        try {
            let ans = createCellStyle(wb);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("createCellStyle")
        }

    pub def createFont(wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.Font, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.createFont(); 
        try {
            let ans = createFont(wb);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("createFont")
        }

    pub def createDataFormat(wb: ##org.apache.poi.ss.usermodel.Workbook): Result[##org.apache.poi.ss.usermodel.DataFormat, String] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.createDataFormat(); 
        try {
            let ans = createDataFormat(wb);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("createDataFormat")
        }

    /// Orphan - DataFormat not used enough to warrant a module.
    pub def getFormat(dfmt: ##org.apache.poi.ss.usermodel.DataFormat,
                      format: String): Result[Int16, String] & Impure = 
        import org.apache.poi.ss.usermodel.DataFormat.getFormat(String); 
        try {
            let ans = getFormat(dfmt, format);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getFormat")
        }

}