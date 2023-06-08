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


mod Interop.Apache.Poi.HSSF.Usermodel {

    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook
    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook.{MutWorkbook}
    use Interop.Apache.Poi.HSSF.Usermodel.HSSFWorkbook
    use Interop.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.{HSSFWorkbook}

    pub enum HSSFWorkbook[_: Region](##org.apache.poi.hssf.usermodel.HSSFWorkbook)


    instance Interop.Apache.Poi.SS.Usermodel.SubMutWorkbook[HSSFWorkbook] {
        pub def toMutWorkbook(d: HSSFWorkbook[r]): MutWorkbook[r] = 
            let HSSFWorkbook(d1) = d;
            MutWorkbook(checked_cast(d1))

        pub def fromMutWorkbook(d: MutWorkbook[r]): HSSFWorkbook[r] = 
            let MutWorkbook(d1) = d;
            HSSFWorkbook(unchecked_cast(d1 as ##org.apache.poi.hssf.usermodel.HSSFWorkbook))
    }

}

mod Interop.Apache.Poi.HSSF.Usermodel.HSSFWorkbook {

    use Interop.Apache.Poi.HSSF.Usermodel.HSSFWorkbook
    use Interop.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.{HSSFWorkbook}
    use Interop.IO.SubInputStream


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
    pub def newWithInputStream[ins: Region -> Type, r: Region](_: Region[r], ins: ins[r]): Result[String, HSSFWorkbook[r]] \ IO with SubInputStream[ins] = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.hssf.usermodel.HSSFWorkbook(##java.io.InputStream): ##org.apache.poi.hssf.usermodel.HSSFWorkbook \ IO as newHSSFWorkbook;
            let ins1 = Interop.IO.SubInputStream.toInputStream(ins) |> Interop.IO.InputStream.getHH;
            newHSSFWorkbook(ins1) |> HSSFWorkbook
        })
/*
    ///
    /// Write workbook.
    ///
    pub def writeHSSFWorkbook(wb: ##org.apache.poi.hssf.usermodel.HSSFWorkbook, 
                              path: String): Result[Unit, String] \ IO = 
        import new java.io.FileOutputStream(String): ##java.io.FileOutputStream \ IO as newFileOutputStream;
        import java.io.FileOutputStream.close(): Unit \ IO;
        import org.apache.poi.hssf.usermodel.HSSFWorkbook.write(##java.io.OutputStream): Unit \ IO; 
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

    pub def close!(wb: HSSFWorkbook[r]): Result[String, Unit] \ r = 
        Result.tryCatch(_ -> {
            import org.apache.poi.hssf.usermodel.HSSFWorkbook.close(): Unit \ r; 
            let HSSFWorkbook(wb1) = wb;
            close(wb1)  
        })


}
