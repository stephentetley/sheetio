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

mod Internal.Apache.Poi.SS.Usermodel {


    pub enum Workbook(##org.apache.poi.ss.usermodel.Workbook)

}


mod Internal.Apache.Poi.SS.Usermodel.Workbook {

    use Internal.Apache.Poi.SS.Usermodel.Workbook
    use Internal.Apache.Poi.SS.Usermodel.Workbook.{Workbook}
    use Internal.Apache.Poi.SS.Usermodel.CellStyle
    use Internal.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Internal.Apache.Poi.SS.Usermodel.DataFormat
    use Internal.Apache.Poi.SS.Usermodel.DataFormat.{DataFormat}
    use Internal.Apache.Poi.SS.Usermodel.Font
    use Internal.Apache.Poi.SS.Usermodel.Font.{Font}
    use Internal.Apache.Poi.SS.Usermodel.Sheet
    use Internal.Apache.Poi.SS.Usermodel.Sheet.{Sheet}

    pub def getHH(wb: Workbook): ##org.apache.poi.ss.usermodel.Workbook =
        let Workbook(wb1) = wb;
        wb1
   
    pub def wrapHH(wb: ##org.apache.poi.ss.usermodel.Workbook): Workbook =
        Workbook(wb)
    

    /// Note - imperative operation, use a phantom region...
    pub def close!(_: Region[r], wb: Workbook): Result[String, Unit] \ r =         
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.close(): Unit \ r;
            let Workbook(wb1) = wb;
            close(wb1)
        })


    pub def getNumberOfSheets(wb: Workbook): Int32 = 
        import org.apache.poi.ss.usermodel.Workbook.getNumberOfSheets(): Int32 \ {};
        let Workbook(wb1) = wb;
        getNumberOfSheets(wb1)

    pub def getSheetName(ix: Int32, wb: Workbook): String = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetName(Int32): String \ {};
        let Workbook(wb1) = wb;
        getSheetName(wb1, ix)


    pub def getSheet(name: String, wb: Workbook): Option[Sheet] = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String): ##org.apache.poi.ss.usermodel.Sheet \ {};
        let Workbook(wb1) = wb;
        getSheet(wb1, name) |> Object.toOption |> Option.map(Sheet)


    pub def getSheetAt(ix: Int32, wb: Workbook): Result[String, Sheet] = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32): ##org.apache.poi.ss.usermodel.Sheet \ {};
            let Workbook(wb1) = wb;
            getSheetAt(wb1, ix) |> Sheet
        })


    // pub def getFontAt(ix: Int32, wb: Workbook[r]): Option[Font[r]] \ r = 
    //     import org.apache.poi.ss.usermodel.Workbook.getFontAt(Int32): ##org.apache.poi.ss.usermodel.Font \ r; 
    //     let Workbook(wb1) = wb;
    //     getFontAt(wb1, ix) |> Object.toOption |> Option.map(Font)


}
