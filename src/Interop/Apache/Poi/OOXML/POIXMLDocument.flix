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

mod Interop.Apache.Poi.OOXML {


    pub enum POIXMLDocument[_: Region](##org.apache.poi.ooxml.POIXMLDocument)


    pub class SubPOIXMLDocument[t: Region -> Type] {
        pub def toPOIXMLDocument[r: Region](s: t[r]): POIXMLDocument[r]
        pub def fromPOIXMLDocument[r: Region](s: POIXMLDocument[r]): t[r]
    }


}

mod Interop.Apache.Poi.OOXML.POIXMLDocument {

    use Interop.Apache.Poi.OOXML.POIXMLDocument
    use Interop.Apache.Poi.OOXML.POIXMLDocument.{POIXMLDocument}
    use Interop.Apache.Poi.OOXML.SubPOIXMLDocument
    use Interop.IO.OutputStream
    use Interop.IO.SubOutputStream
    

    // XSSF (OOXML) concretely

    pub def getHH(wb: POIXMLDocument[r]): ##org.apache.poi.ooxml.POIXMLDocument =
        let POIXMLDocument(wb1) = wb;
        wb1

    pub def wrapHH(_: Region[r], wb: ##org.apache.poi.ooxml.POIXMLDocument): POIXMLDocument[r] =
        POIXMLDocument(wb)



    ///
    /// Write workbook.
    ///
    pub def write!(out: OutputStream[r], wb: POIXMLDocument[r]): Result[String, Unit] \ {IO, r} = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ooxml.POIXMLDocument.write(##java.io.OutputStream): Unit \ {IO, r};
            let POIXMLDocument(wb1) = wb;
            let out1 = Interop.IO.OutputStream.getHH(out);
            write(wb1, out1)
        })


}
