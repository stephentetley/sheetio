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

namespace SheetReader/Internal/Prim {


    pub def newXSSFWorkbook(path: String): Result[##org.apache.poi.xssf.usermodel.XSSFWorkbook, System.Error] & Impure = 
        import new java.io.File(String) as newFile;
        import new java.io.FileInputStream(##java.io.File) as newFileInputStream;
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.InputStream) as newXSSFWorkbook;
        try {
            let file1 = newFile(path);
            let fistream = newFileInputStream(file1);
            let istream = fistream as ##java.io.InputStream;
            let wb = newXSSFWorkbook(istream);
            Ok(wb)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("newXSSFWorkbook")) 
        }

    pub def closeXSSFWorkbook(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFWorkbook.close(); 
        try {
            let _ = wb.close();
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err(Generic("closeXSSFWorkbook")) 
        }


    pub def getNumberOfSheets(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): Result[Int32, System.Error] & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFWorkbook.getNumberOfSheets(); 
        try {
            let ans = wb.getNumberOfSheets();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getNumberOfSheets")) 
        }

    pub def getSheetName(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook, ix: Int32): Result[String, System.Error] & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFWorkbook.getSheetName(Int32); 
        try {
            let ans = wb.getSheetName(ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getSheetName")) 
        }

}