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

mod Interop.Apache.Poi.XSSF.Usermodel {

    use Interop.Apache.Poi.OOXML.POIXMLDocument
    use Interop.Apache.Poi.OOXML.POIXMLDocument.{POIXMLDocument}
    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook
    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook.{MutWorkbook}
    use Interop.Apache.Poi.XSSF.Usermodel.XSSFWorkbook
    use Interop.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.{XSSFWorkbook}

    pub enum XSSFWorkbook[_: Region](##org.apache.poi.xssf.usermodel.XSSFWorkbook)

    instance Interop.Apache.Poi.OOXML.SubPOIXMLDocument[XSSFWorkbook] {
        pub def toPOIXMLDocument(d: XSSFWorkbook[r]): POIXMLDocument[r] = 
            let XSSFWorkbook(d1) = d;
            POIXMLDocument(checked_cast(d1))

        pub def fromPOIXMLDocument(d: POIXMLDocument[r]): XSSFWorkbook[r] = 
            let POIXMLDocument(d1) = d;
            XSSFWorkbook(unchecked_cast(d1 as ##org.apache.poi.xssf.usermodel.XSSFWorkbook))
    }

    instance Interop.Apache.Poi.SS.Usermodel.SubMutWorkbook[XSSFWorkbook] {
        pub def toMutWorkbook(d: XSSFWorkbook[r]): MutWorkbook[r] = 
            let XSSFWorkbook(d1) = d;
            MutWorkbook(checked_cast(d1))

        pub def fromMutWorkbook(d: MutWorkbook[r]): XSSFWorkbook[r] = 
            let MutWorkbook(d1) = d;
            XSSFWorkbook(unchecked_cast(d1 as ##org.apache.poi.xssf.usermodel.XSSFWorkbook))
    }

}

mod Interop.Apache.Poi.XSSF.Usermodel.XSSFWorkbook {
    
    use Interop.IO.File
    use Interop.IO.InputStream
    use Interop.IO.SubInputStream

    use Interop.Apache.Poi.XSSF.Usermodel.XSSFWorkbook
    use Interop.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.{XSSFWorkbook}
    

    // XSSF (OOXML) concretely

    pub def getHH(wb: XSSFWorkbook[r]): ##org.apache.poi.xssf.usermodel.XSSFWorkbook =
        let XSSFWorkbook(wb1) = wb;
        wb1

    pub def wrapHH(_: Region[r], wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): XSSFWorkbook[r] =
        XSSFWorkbook(wb)

    ///
    /// Create new for writing.
    ///
    pub def new(_: Region[r]): XSSFWorkbook[r] \ r  = 
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook():##org.apache.poi.xssf.usermodel.XSSFWorkbook \ r as newXSSFWorkbook;
        newXSSFWorkbook() |> XSSFWorkbook


    ///
    /// Open for reading.
    ///
    pub def newWithFile(_: Region[r], file: File[r1]): Result[String, XSSFWorkbook[r]] \ {IO, r, r1} = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.File): ##org.apache.poi.xssf.usermodel.XSSFWorkbook \ {IO, r, r1} as newXSSFWorkbook;
            let file1 = Interop.IO.File.getHH(file);
            newXSSFWorkbook(file1) |> XSSFWorkbook
        })


    ///
    /// Open for reading.
    ///
    pub def newWithInputStream(_: Region[r], ins: InputStream[r1]): Result[String, XSSFWorkbook[r]] \ {IO, r, r1}  = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.InputStream): ##org.apache.poi.xssf.usermodel.XSSFWorkbook \ {IO, r, r1} as newXSSFWorkbook;
            let ins1 = Interop.IO.InputStream.getHH(ins);
            newXSSFWorkbook(ins1) |> XSSFWorkbook
        })


    /// `write(OutputStream)` is method on superclass `POIXMLDocument` 


    pub def close!(wb: XSSFWorkbook[r]): Result[String, Unit] \ {IO, r} = 
        Result.tryCatch(_ -> {
            import org.apache.poi.xssf.usermodel.XSSFWorkbook.close(): Unit \ {IO, r}; 
            let XSSFWorkbook(wb1) = wb;
            close(wb1)
        })

}
