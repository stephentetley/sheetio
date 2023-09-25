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

mod Interop.Apache.Poi.SS.Usermodel {

    /// Mutable / Interface

    pub enum MutWorkbook[_: Region](##org.apache.poi.ss.usermodel.Workbook)


    pub class SubMutWorkbook[t: Region -> Type] {
        pub def toMutWorkbook(o: t[r]): MutWorkbook[r]
        pub def fromMutWorkbook(w: MutWorkbook[r]): t[r]
    }

}


mod Interop.Apache.Poi.SS.Usermodel.MutWorkbook {

    use Basis.IO.OutputStream

    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook
    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook.{MutWorkbook}
    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Interop.Apache.Poi.SS.Usermodel.DataFormat
    use Interop.Apache.Poi.SS.Usermodel.DataFormat.{DataFormat}
    use Interop.Apache.Poi.SS.Usermodel.Font
    use Interop.Apache.Poi.SS.Usermodel.Font.{Font}
    use Interop.Apache.Poi.SS.Usermodel.MutSheet
    use Interop.Apache.Poi.SS.Usermodel.MutSheet.{MutSheet}

    pub def getHH(wb: MutWorkbook[r]): ##org.apache.poi.ss.usermodel.Workbook =
        let MutWorkbook(wb1) = wb;
        wb1
   
    pub def wrapHH(_: Region[r], s: ##org.apache.poi.ss.usermodel.Workbook): MutWorkbook[r] =
        MutWorkbook(s)
        
    pub def write!(out: OutputStream[r1], wb: MutWorkbook[r]): Result[String, Unit] \ r =         
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.write(##java.io.OutputStream): Unit \ r;
            let MutWorkbook(wb1) = wb;
            let out1 = Interop.IO.OutputStream.getHH(out);
            write(wb1, out1)
        })

    pub def close!(wb: MutWorkbook[r]): Result[String, Unit] \ r =         
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.close(): Unit \ r;
            let MutWorkbook(wb1) = wb;
            close(wb1)
        })

    pub def createSheet!(name: String, wb: MutWorkbook[r]): Result[String, MutSheet[r]] \ r =
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.createSheet(String): ##org.apache.poi.ss.usermodel.Sheet \ r;
            let MutWorkbook(wb1) = wb;
            createSheet(wb1, name) |> MutSheet
        })

    pub def cloneSheet!(num: Int32, wb: MutWorkbook[r]): Result[String, MutSheet[r]] \ r =
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.cloneSheet(Int32): ##org.apache.poi.ss.usermodel.Sheet \ r;
            let MutWorkbook(wb1) = wb;
            cloneSheet(wb1, num) |> MutSheet
        })

    pub def getNumberOfSheets(wb: MutWorkbook[r]): Int32 \ r = 
        import org.apache.poi.ss.usermodel.Workbook.getNumberOfSheets(): Int32 \ r;
        let MutWorkbook(wb1) = wb;
        getNumberOfSheets(wb1)

    pub def getSheetName(ix: Int32, wb: MutWorkbook[r]): String \ r = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetName(Int32): String \ r;
        let MutWorkbook(wb1) = wb;
        getSheetName(wb1, ix)


    pub def getSheet(name: String, wb: MutWorkbook[r]): Option[MutSheet[r]] \ r = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String): ##org.apache.poi.ss.usermodel.Sheet \ r;
        let MutWorkbook(wb1) = wb;
        getSheet(wb1, name) |> Object.toOption |> Option.map(MutSheet)


    pub def getSheetAt(ix: Int32, wb: MutWorkbook[r]): Result[String, MutSheet[r]] \ r = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32): ##org.apache.poi.ss.usermodel.Sheet \ r;
            let MutWorkbook(wb1) = wb;
            getSheetAt(wb1, ix) |> MutSheet
        })


    pub def getFontAt(ix: Int32, wb: MutWorkbook[r]): Option[Font[r]] \ r = 
        import org.apache.poi.ss.usermodel.Workbook.getFontAt(Int32): ##org.apache.poi.ss.usermodel.Font \ r; 
        let MutWorkbook(wb1) = wb;
        getFontAt(wb1, ix) |> Object.toOption |> Option.map(Font)

    /// Updates the workbook's style table.
    pub def createCellStyle!(wb: MutWorkbook[r]): Result[String, CellStyle[r]] \ r = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.createCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ r; 
            let MutWorkbook(wb1) = wb;
            createCellStyle(wb1) |> CellStyle
        })

    pub def createFont!(wb: MutWorkbook[r]): Font[r] \ r = 
        import org.apache.poi.ss.usermodel.Workbook.createFont(): ##org.apache.poi.ss.usermodel.Font \ r; 
        let MutWorkbook(wb1) = wb;
        createFont(wb1) |> Font

    pub def createDataFormat!(wb: MutWorkbook[r]): DataFormat[r] \ r = 
        import org.apache.poi.ss.usermodel.Workbook.createDataFormat(): ##org.apache.poi.ss.usermodel.DataFormat \ r; 
        let MutWorkbook(wb1) = wb;
        createDataFormat(wb1) |> DataFormat

}
