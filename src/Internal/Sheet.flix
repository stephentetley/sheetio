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

    pub def getSheetName(sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[String, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getSheetName(); 
        try {
            let ans = getSheetName(sheet);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getSheetName [Sheet]")
        }

    pub def getFirstRowNum(sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getFirstRowNum(); 
        try {
            let ans = getFirstRowNum(sheet);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getFirstRowNum")
        }

    pub def getLastRowNum(sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getLastRowNum(); 
        try {
            let ans = getLastRowNum(sheet);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getLastRowNum")
        }

    pub def getPhysicalNumberOfRows(sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getPhysicalNumberOfRows(); 
        try {
            let ans = getPhysicalNumberOfRows(sheet);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getPhysicalNumberOfRows")
        }

    pub def getWorkbook(sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[##org.apache.poi.ss.usermodel.Workbook, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getWorkbook(); 
        try {
            let ans = getWorkbook(sheet);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getWorkbook [Sheet]")
        }


    pub def getRow(sheet: ##org.apache.poi.ss.usermodel.Sheet,
                   ix: Int32): Result[##org.apache.poi.ss.usermodel.Row, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getRow(Int32); 
        try {
            let ans = getRow(sheet, ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getRow")
        }

    pub def createRow(sheet: ##org.apache.poi.ss.usermodel.Sheet,
                      ix: Int32): Result[##org.apache.poi.ss.usermodel.Row, String] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.createRow(Int32); 
        try {
            let ans = createRow(sheet, ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("createRow")
        }

}
