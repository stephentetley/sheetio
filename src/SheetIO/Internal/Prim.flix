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


namespace SheetIO/Internal/Prim {


    use SheetIO/Internal/PrimError.{PrimError};
    
    // XSSF (OOXML) concretely

    ///
    /// Create new for writing.
    ///
    pub def newXSSFWorkbook(r: Region[r]): Result[##org.apache.poi.xssf.usermodel.XSSFWorkbook, PrimError] \ { Impure, Write(r) } = 
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook():##org.apache.poi.xssf.usermodel.XSSFWorkbook & Impure as newXSSFWorkbook;
        try {
            let wb = newXSSFWorkbook();
            Ok(wb)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "newXSSFWorkbook", ex) 
        }


    ///
    /// Open for reading.
    ///
    pub def openXSSFWorkbook(path: String): Result[##org.apache.poi.xssf.usermodel.XSSFWorkbook, String] & Impure = 
        import new java.io.File(String): ##java.io.File & Impure as newFile;
        import new java.io.FileInputStream(##java.io.File): ##java.io.FileInputStream & Impure as newFileInputStream;
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.InputStream): ##org.apache.poi.xssf.usermodel.XSSFWorkbook & Impure as newXSSFWorkbook;
        try {
            let file1 = newFile(path);
            let fistream = newFileInputStream(file1);
            let istream = fistream as ##java.io.InputStream;
            let wb = newXSSFWorkbook(istream);
            Ok(wb)
        } catch {
            case _: ##java.lang.Exception => Err("openXSSFWorkbook")
        }


    ///
    /// Write workbook.
    ///
    pub def writeXSSFWorkbook(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook, 
                              path: String): Result[Unit, String] & Impure = 
        import new java.io.FileOutputStream(String): ##java.io.FileOutputStream & Impure as newFileOutputStream;
        import java.io.FileOutputStream.close(): Unit & Impure;
        import org.apache.poi.ooxml.POIXMLDocument.write(##java.io.OutputStream): Unit & Impure;
        try {
            let fostream = newFileOutputStream(path);
            let fostream1 = fostream as ##java.io.OutputStream;
            let wb1 = wb as ##org.apache.poi.ooxml.POIXMLDocument;
            write(wb1, fostream1);
            close(fostream);
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err("writeXSSFWorkbook")
        }


    pub def closeXSSFWorkbook(wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): Result[Unit, String] & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFWorkbook.close(): Unit & Impure; 
        try {
            let _ = close(wb);
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err("closeXSSFWorkbook")
        }

    ///
    /// Create new for writing.
    ///
    pub def newHSSFWorkbook(): Result[##org.apache.poi.hssf.usermodel.HSSFWorkbook, String] & Impure = 
        import new org.apache.poi.hssf.usermodel.HSSFWorkbook(): ##org.apache.poi.hssf.usermodel.HSSFWorkbook & Impure as newHSSFWorkbook;
        try {
            let wb = newHSSFWorkbook();
            Ok(wb)
        } catch {
            case _: ##java.lang.Exception => Err("newHSSFWorkbook")
        }

    ///
    /// Open for reading.
    ///
    pub def openHSSFWorkbook(path: String): Result[##org.apache.poi.hssf.usermodel.HSSFWorkbook, String] & Impure = 
        import new java.io.File(String): ##java.io.File & Impure as newFile;
        import new java.io.FileInputStream(##java.io.File): ##java.io.FileInputStream & Impure as newFileInputStream;
        import new org.apache.poi.hssf.usermodel.HSSFWorkbook(##java.io.InputStream): ##org.apache.poi.hssf.usermodel.HSSFWorkbook & Impure as newHSSFWorkbook;
        try {
            let file1 = newFile(path);
            let fistream = newFileInputStream(file1);
            let istream = fistream as ##java.io.InputStream;
            let wb = newHSSFWorkbook(istream);
            Ok(wb)
        } catch {
            case _: ##java.lang.Exception => Err("openHSSFWorkbook")
        }

    ///
    /// Write workbook.
    ///
    pub def writeHSSFWorkbook(wb: ##org.apache.poi.hssf.usermodel.HSSFWorkbook, 
                              path: String): Result[Unit, String] & Impure = 
        import new java.io.FileOutputStream(String): ##java.io.FileOutputStream & Impure as newFileOutputStream;
        import java.io.FileOutputStream.close(): Unit & Impure;
        import org.apache.poi.hssf.usermodel.HSSFWorkbook.write(##java.io.OutputStream): Unit & Impure; 
        try {
            let fostream = newFileOutputStream(path);
            let fostream1 = fostream as ##java.io.OutputStream;
            write(wb, fostream1);
            close(fostream);
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err("writeHSSFWorkbook")
        }


    pub def closeHSSFWorkbook(wb: ##org.apache.poi.hssf.usermodel.HSSFWorkbook): Result[Unit, String] & Impure = 
        import org.apache.poi.hssf.usermodel.HSSFWorkbook.close(): Unit & Impure; 
        try {
            let _ = close(wb);
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err("closeHSSFWorkbook")
        }


}
