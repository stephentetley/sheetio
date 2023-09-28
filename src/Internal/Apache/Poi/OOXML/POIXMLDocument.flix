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

mod Internal.Apache.Poi.OOXML {


    pub enum POIXMLDocument[_: Region](##org.apache.poi.ooxml.POIXMLDocument)

}

mod Internal.Apache.Poi.OOXML.POIXMLDocument {

    use Basis.IO.OutputStream
    use Basis.IO.OutputStream.{OutputStream}

    use Internal.Apache.Poi.OOXML.POIXMLDocument
    use Internal.Apache.Poi.OOXML.POIXMLDocument.{POIXMLDocument}
    
    

    // XSSF (OOXML) concretely

    ///
    /// Write document.
    ///
    pub def write!(out: OutputStream[r], wb: POIXMLDocument[r1]): Result[String, Unit] \ {r, r1} = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ooxml.POIXMLDocument.write(##java.io.OutputStream): Unit \ {r, r1};
            let POIXMLDocument(wb1) = wb;
            let OutputStream(out1) = out;
            write(wb1, out1)
        })


}
