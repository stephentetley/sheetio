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


namespace Interop/Apache/Poi/HSSF/Usermodel {

    use Interop/Apache/Poi/SS/Usermodel.Workbook;
    use Interop/Apache/Poi/SS/Usermodel.Workbook.{Workbook};

    pub enum HSSFWorkbook[_: Region](##org.apache.poi.hssf.usermodel.HSSFWorkbook)

    instance Scoped[HSSFWorkbook] {
        pub def regionOf(_: HSSFWorkbook[r]): Region[r] = () as Region[r]
    }

    instance Interop/Apache/Poi/SS/Usermodel.SubWorkbook[HSSFWorkbook] {
        pub def toWorkbook(d: HSSFWorkbook[r]): Workbook[r] = 
            let HSSFWorkbook(d1) = d;
            Workbook(d1 as ##org.apache.poi.ss.usermodel.Workbook)

        pub def fromWorkbook(d: Workbook[r]): HSSFWorkbook[r] = 
            let Workbook(d1) = d;
            HSSFWorkbook(d1 as ##org.apache.poi.hssf.usermodel.HSSFWorkbook) 
    }

}

namespace Interop/Apache/Poi/HSSF/Usermodel/HSSFWorkbook {

    use Interop/Apache/Poi/HSSF/Usermodel.HSSFWorkbook;
    use Interop/Apache/Poi/HSSF/Usermodel.HSSFWorkbook.{HSSFWorkbook};
    use Interop/IO.SubInputStream;


    pub def getHH(wb: HSSFWorkbook[r]): ##org.apache.poi.hssf.usermodel.HSSFWorkbook =
        let HSSFWorkbook(wb1) = wb;
        wb1

    pub def wrapHH(_: Region[r], wb: ##org.apache.poi.hssf.usermodel.HSSFWorkbook): HSSFWorkbook[r] =
        HSSFWorkbook(wb)


    ///
    /// Create new for writing.
    ///
    pub def new(_: Region[r]): HSSFWorkbook[r] \ IO =
        import new org.apache.poi.hssf.usermodel.HSSFWorkbook(): ##org.apache.poi.hssf.usermodel.HSSFWorkbook \ IO as newHSSFWorkbook;
        newHSSFWorkbook() |> HSSFWorkbook
        
    ///
    /// Open for reading.
    ///
    pub def newWithInputStream[ins: Region -> Type, r: Region](_: Region[r], ins: ins[r]): Result[HSSFWorkbook[r], String] \ IO with SubInputStream[ins] = 
        try {
            import new org.apache.poi.hssf.usermodel.HSSFWorkbook(##java.io.InputStream): ##org.apache.poi.hssf.usermodel.HSSFWorkbook \ IO as newHSSFWorkbook;
            let ins1 = Interop/IO/SubInputStream.toInputStream(ins) |> Interop/IO/InputStream.getHH;
            newHSSFWorkbook(ins1) |> HSSFWorkbook |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }
/*
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
*/

    pub def close!(wb: HSSFWorkbook[r]): Result[Unit, String] \ {Write(r)} = 
        try {
            import org.apache.poi.hssf.usermodel.HSSFWorkbook.close(): Unit \ r; 
            let HSSFWorkbook(wb1) = wb;
            let _ = close(wb1);
            Ok()
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


}
