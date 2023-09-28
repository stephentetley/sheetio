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


mod SheetIO.Internal.Apache.Poi.HSSF.Usermodel {


    pub enum HSSFWorkbook[_: Region](##org.apache.poi.hssf.usermodel.HSSFWorkbook)

}

mod SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook {
    
    use Basis.IO.InputStream
    use Basis.IO.InputStream.{InputStream}
    use Basis.IO.OutputStream
    use Basis.IO.OutputStream.{OutputStream}

    use SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook
    use SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.{HSSFWorkbook}
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.{MutWorkbook}


    ///
    /// Create new for writing.
    ///
    pub def new(_: Region[r]): HSSFWorkbook[r] \ IO =
        import new org.apache.poi.hssf.usermodel.HSSFWorkbook(): ##org.apache.poi.hssf.usermodel.HSSFWorkbook \ IO as newHSSFWorkbook;
        newHSSFWorkbook() |> HSSFWorkbook
        
    ///
    /// Open for reading.
    ///
    pub def newWithInputStream(_: Region[r], ins: InputStream[r]): Result[String, HSSFWorkbook[r]] \ IO  = 
        Result.tryCatch(_ -> {
            import new org.apache.poi.hssf.usermodel.HSSFWorkbook(##java.io.InputStream): ##org.apache.poi.hssf.usermodel.HSSFWorkbook \ IO as newHSSFWorkbook;
            let InputStream(ins1) = ins;
            newHSSFWorkbook(ins1) |> HSSFWorkbook
        })

    pub def asMutWorkbook(d: HSSFWorkbook[r]): MutWorkbook[r] = 
        let HSSFWorkbook(d1) = d;
        MutWorkbook(checked_cast(d1))
        
    ///
    /// Write workbook.
    ///
    pub def write!(out: OutputStream[r1], wb: HSSFWorkbook[r]): Result[String, Unit] \ {IO, r, r1} = 
        Result.tryCatch(_ -> {
            import org.apache.poi.hssf.usermodel.HSSFWorkbook.write(##java.io.OutputStream): Unit \ {IO, r, r1}; 
            let HSSFWorkbook(wb1) = wb;
            let OutputStream(out1) = out;
            write(wb1, out1)
        })


    pub def close!(wb: HSSFWorkbook[r]): Result[String, Unit] \ r = 
        Result.tryCatch(_ -> {
            import org.apache.poi.hssf.usermodel.HSSFWorkbook.close(): Unit \ r; 
            let HSSFWorkbook(wb1) = wb;
            close(wb1)  
        })


}
