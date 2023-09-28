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

mod Internal.Apache.Poi.XSSF.Usermodel {

    pub enum XSSFWorkbook[_: Region](##org.apache.poi.xssf.usermodel.XSSFWorkbook)


}

mod Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook {
    
    use Basis.IO.File
    use Basis.IO.File.{File}
    use Basis.IO.InputStream
    use Basis.IO.InputStream.{InputStream}

    use Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook
    use Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.{XSSFWorkbook}
    use Internal.Apache.Poi.SS.Usermodel.MutWorkbook
    use Internal.Apache.Poi.SS.Usermodel.MutWorkbook.{MutWorkbook}
    use Internal.Apache.Poi.OOXML.POIXMLDocument
    use Internal.Apache.Poi.OOXML.POIXMLDocument.{POIXMLDocument}


    // XSSF (OOXML) concretely


    ///
    /// Create new for writing.
    ///
    pub def new(_: Region[r]): XSSFWorkbook[r] \ r  = 
        import new org.apache.poi.xssf.usermodel.XSSFWorkbook():##org.apache.poi.xssf.usermodel.XSSFWorkbook \ r as newXSSFWorkbook;
        newXSSFWorkbook() |> XSSFWorkbook

    pub def asMutWorkbook(d: XSSFWorkbook[r]): MutWorkbook[r] = 
        let XSSFWorkbook(d1) = d;
        MutWorkbook(checked_cast(d1))

    pub def asPOIXMLDocument(d: XSSFWorkbook[r]): POIXMLDocument[r] = 
        let XSSFWorkbook(d1) = d;
        POIXMLDocument(checked_cast(d1))

    ///
    ///
    pub def newWithFile(_: Region[r], file: File): Result[String, XSSFWorkbook[r]] \ {IO, r} = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.File): ##org.apache.poi.xssf.usermodel.XSSFWorkbook \ {IO, r} as newXSSFWorkbook;
            let File(file1) = file;
            newXSSFWorkbook(file1) |> XSSFWorkbook
        })


    ///
    /// Open for reading.
    ///
    pub def newWithInputStream(_: Region[r], ins: InputStream[r1]): Result[String, XSSFWorkbook[r]] \ {IO, r, r1}  = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.xssf.usermodel.XSSFWorkbook(##java.io.InputStream): ##org.apache.poi.xssf.usermodel.XSSFWorkbook \ {IO, r, r1} as newXSSFWorkbook;
            let InputStream(ins1) = ins;
            newXSSFWorkbook(ins1) |> XSSFWorkbook
        })


    ///
    ///
    pub def newWithPathString(_: Region[r], path: String): Result[String, XSSFWorkbook[r]] \ {IO, r} = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.xssf.usermodel.XSSFWorkbook(String): ##org.apache.poi.xssf.usermodel.XSSFWorkbook \ {IO, r} as new1;
            new1(path) |> XSSFWorkbook
        })


    /// `write(OutputStream)` is method on superclass `POIXMLDocument` 


    pub def close!(wb: XSSFWorkbook[r]): Result[String, Unit] \ {IO, r} = 
        Result.tryCatch(_ -> {
            import org.apache.poi.xssf.usermodel.XSSFWorkbook.close(): Unit \ {IO, r}; 
            let XSSFWorkbook(wb1) = wb;
            close(wb1)
        })

}
