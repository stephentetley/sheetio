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

namespace Interop/Apache/Poi/OOXML {


    pub enum POIXMLDocument[_: Region](##org.apache.poi.ooxml.POIXMLDocument)

    instance Scoped[POIXMLDocument] {
        pub def regionOf(_: POIXMLDocument[r]): Region[r] = unsafe_cast () as Region[r]
    }

    pub class SubPOIXMLDocument[t: Region -> Type] {
        pub def toPOIXMLDocument[r: Region](s: t[r]): POIXMLDocument[r]
        pub def fromPOIXMLDocument[r: Region](s: POIXMLDocument[r]): t[r]
    }


}

namespace Interop/Apache/Poi/OOXML/POIXMLDocument {

    use Interop/Apache/Poi/OOXML.POIXMLDocument;
    use Interop/Apache/Poi/OOXML.POIXMLDocument.{POIXMLDocument};
    use Interop/Apache/Poi/OOXML.SubPOIXMLDocument;
    use Interop/IO.OutputStream;
    use Interop/IO.SubOutputStream;
    

    // XSSF (OOXML) concretely

    pub def getHH(wb: POIXMLDocument[r]): ##org.apache.poi.ooxml.POIXMLDocument =
        let POIXMLDocument(wb1) = wb;
        wb1

    pub def wrapHH(_: Region[r], wb: ##org.apache.poi.ooxml.POIXMLDocument): POIXMLDocument[r] =
        POIXMLDocument(wb)



    ///
    /// Write workbook.
    ///
    pub def write![r: Region, outs: Region -> Type, wb: Region -> Type]
                (out: outs[r], wb: wb[r]): Result[Unit, String] \ {IO, Read(r)} with SubPOIXMLDocument[wb], SubOutputStream[outs] = 
        try {
            import org.apache.poi.ooxml.POIXMLDocument.write(##java.io.OutputStream): Unit \ {IO, r};
            let wb1  = Interop/Apache/Poi/OOXML/SubPOIXMLDocument.toPOIXMLDocument(wb) |> getHH;
            let out1 = Interop/IO/SubOutputStream.toOutputStream(out) |> Interop/IO/OutputStream.getHH;
            write(wb1, out1);
            Ok()
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


}
