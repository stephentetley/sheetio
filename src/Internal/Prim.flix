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


    ///
    /// Create new for writing.
    ///
    pub def newXSSFWorkbook(): Result[##org.apache.poi.xssf.usermodel.XSSFWorkbook, System.Error] & Impure = 
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook() as newXSSFWorkbook;
        try {
            let wb = newXSSFWorkbook();
            Ok(wb)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("newXSSFWorkbook")) 
        }


    ///
    /// Open for reading.
    ///
    pub def openXSSFWorkbook(path: String): Result[##org.apache.poi.xssf.usermodel.XSSFWorkbook, System.Error] & Impure = 
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
            case _: ##java.lang.Exception => Err(Generic("openXSSFWorkbook")) 
        }


    ///
    /// Open for reading.
    ///
    pub def writeXSSFWorkbook(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook, 
                              path: String): Result[Unit, System.Error] & Impure = 
        import new java.io.FileOutputStream(String) as newFileOutputStream;
        import java.io.FileOutputStream.close();
        import org.apache.poi.ooxml.POIXMLDocument.write(##java.io.OutputStream);
        try {
            let fostream = newFileOutputStream(path);
            let fostream1 = fostream as ##java.io.OutputStream;
            let wb1 = wb as ##org.apache.poi.ooxml.POIXMLDocument;
            wb1.write(fostream1);
            fostream.close();
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err(Generic("writeXSSFWorkbook")) 
        }


    pub def closeXSSFWorkbook(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFWorkbook.close(); 
        try {
            let _ = wb.close();
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err(Generic("closeXSSFWorkbook")) 
        }


    pub def createSheet(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                        name: String): Result[##org.apache.poi.ss.usermodel.Sheet, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.createSheet(String); 
        try {
            let ans = wb.createSheet(name);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("createSheet")) 
        }


    pub def getNumberOfSheets(wb: ##org.apache.poi.ss.usermodel.Workbook): Result[Int32, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getNumberOfSheets(); 
        try {
            let ans = wb.getNumberOfSheets();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getNumberOfSheets")) 
        }

    pub def getSheetName(wb: ##org.apache.poi.ss.usermodel.Workbook, ix: Int32): Result[String, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetName(Int32); 
        try {
            let ans = wb.getSheetName(ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getSheetName")) 
        }

    pub def getSheet(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                     name: String): Result[##org.apache.poi.ss.usermodel.Sheet, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String); 
        try {
            let ans = wb.getSheet(name);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getSheet")) 
        }

    pub def getSheetAt(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                       ix: Int32): Result[##org.apache.poi.ss.usermodel.Sheet, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32); 
        try {
            let ans = wb.getSheetAt(ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getSheetAt")) 
        }        

}