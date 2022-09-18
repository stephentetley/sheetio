/*
 * Copyright 2022 Stephen Tetley
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

namespace Interop/Apache/Poi/XSSF/Usermodel {


    pub enum XSSFWorkbook[_: Region](##org.apache.poi.xssf.usermodel.XSSFWorkbook)

    instance Scoped[XSSFWorkbook] {
        pub def regionOf(_: XSSFWorkbook[r]): Region[r] = () as Region[r]
    }
}

namespace Interop/Apache/Poi/XSSF/Usermodel/XSSFWorkbook {

    use Interop/Apache/Poi/XSSF/Usermodel.XSSFWorkbook;
    use Interop/Apache/Poi/XSSF/Usermodel.XSSFWorkbook.{XSSFWorkbook};
    use Interop/IO.InputStream;
    
    

    // XSSF (OOXML) concretely

    pub def getHH(wb: XSSFWorkbook[r]): ##org.apache.poi.xssf.usermodel.XSSFWorkbook =
        let XSSFWorkbook(wb1) = wb;
        wb1

    pub def wrapHH(_: Region[r], wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): XSSFWorkbook[r] =
        XSSFWorkbook(wb)

    ///
    /// Create new for writing.
    ///
    pub def new(r: Region[r]): XSSFWorkbook[r] \ Write(r)  = 
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook():##org.apache.poi.xssf.usermodel.XSSFWorkbook \ r as newXSSFWorkbook;
        newXSSFWorkbook() |> XSSFWorkbook

    ///
    /// Open for reading.
    ///
    pub def newWithInputStream(ins: InputStream[r]): Result[XSSFWorkbook[r], String] \ IO = 
        try {
            import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.InputStream): ##org.apache.poi.xssf.usermodel.XSSFWorkbook & Impure as newXSSFWorkbook;
            let ins1 = Interop/IO/InputStream.getHH(ins);
            newXSSFWorkbook(ins1) |> XSSFWorkbook |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


    /// `write(OutputStream)` and `close()` are methods on superclass `POIXMLDocument` 

/*
    ///
    /// Write workbook.
    ///
    pub def writeToFile(out: File[r], wb: XSSFWorkbook[r]): Result[Unit, String] & Impure = 
        import org.apache.poi.ooxml.POIXMLDocument.write(##java.io.File): Unit & Impure;
        try {
            let XSSFWorkbook(wb1) = wb;
            let out1 = Interop/IO/File.getHH(out);
            write(wb1, out1);
            Ok()
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


    pub def closeXSSFWorkbook(wb: XSSFWorkbook[r]): Result[Unit, String] \ {Write(r)} = 
        try {
            import org.apache.poi.xssf.usermodel.XSSFWorkbook.close(): Unit \ r; 
            let XSSFWorkbook(wb1) = wb;
            let _ = close(wb1);
            Ok()
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }
*/


}
