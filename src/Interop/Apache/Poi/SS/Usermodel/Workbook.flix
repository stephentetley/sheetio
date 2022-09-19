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

namespace Interop/Apache/Poi/SS/Usermodel {

    /// Mutable / Interface

    pub enum Workbook[_: Region](##org.apache.poi.ss.usermodel.Workbook)

    instance Scoped[Workbook] {
        pub def regionOf(_: Workbook[r]): Region[r] = () as Region[r]
    }

    pub class SubWorkbook[t: Region -> Type] {
        pub def toWorkbook[r: Region](s: t[r]): Workbook[r]
        pub def fromWorkbook[r: Region](s: Workbook[r]): t[r]
    }
}


namespace Interop/Apache/Poi/SS/Usermodel/Workbook {

    use Interop/Apache/Poi/SS/Usermodel.Workbook;
    use Interop/Apache/Poi/SS/Usermodel.Workbook.{Workbook};
    use Interop/Apache/Poi/SS/Usermodel.CellStyle;
    use Interop/Apache/Poi/SS/Usermodel.CellStyle.{CellStyle};
    use Interop/Apache/Poi/SS/Usermodel.DataFormat;
    use Interop/Apache/Poi/SS/Usermodel.DataFormat.{DataFormat};
    use Interop/Apache/Poi/SS/Usermodel.Font;
    use Interop/Apache/Poi/SS/Usermodel.Font.{Font};
    use Interop/Apache/Poi/SS/Usermodel.Sheet;
    use Interop/Apache/Poi/SS/Usermodel.Sheet.{Sheet};

    pub def getHH(wb: Workbook[r]): ##org.apache.poi.ss.usermodel.Workbook =
        let Workbook(wb1) = wb;
        wb1
   
    pub def wrapHH(_: Region[r], s: ##org.apache.poi.ss.usermodel.Workbook): Workbook[r] =
        Workbook(s)
    
    pub def close!(wb: Workbook[r]): Result[Unit, String] \ Write(r) =         
        try {
            import org.apache.poi.ss.usermodel.Workbook.close(): Unit \ r;
            let Workbook(wb1) = wb;
            close(wb1) |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }

    pub def createSheet!(name: String, wb: Workbook[r]): Result[Sheet[r], String] \ Write(r) =         
        try {
            import org.apache.poi.ss.usermodel.Workbook.createSheet(String): ##org.apache.poi.ss.usermodel.Sheet \ r;
            let Workbook(wb1) = wb;
            createSheet(wb1, name) |> Sheet |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


    pub def getNumberOfSheets(wb: Workbook[r]): Int32 \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getNumberOfSheets(): Int32 \ r;
        let Workbook(wb1) = wb;
        getNumberOfSheets(wb1)

    pub def getSheetName(ix: Int32, wb: Workbook[r]): String \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetName(Int32): String \ r;
        let Workbook(wb1) = wb;
        getSheetName(wb1, ix)


    pub def getSheet(name: String, wb: Workbook[r]): Option[Sheet[r]] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String): ##org.apache.poi.ss.usermodel.Sheet \ r;
        let Workbook(wb1) = wb;
        getSheet(wb1, name) |> Object.toOption |> Option.map(Sheet)


    pub def getSheetAt(ix: Int32, wb: Workbook[r]): Result[Sheet[r], String] \ Read(r) = 
        try {
            import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32): ##org.apache.poi.ss.usermodel.Sheet \ r;
            let Workbook(wb1) = wb;
            getSheetAt(wb1, ix) |> Sheet |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


    pub def getFontAt(ix: Int32, wb: Workbook[r]): Option[Font[r]] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getFontAt(Int32): ##org.apache.poi.ss.usermodel.Font \ r; 
        let Workbook(wb1) = wb;
        getFontAt(wb1, ix) |> Object.toOption |> Option.map(Font)


    pub def createCellStyle!(wb: Workbook[r]): Result[CellStyle[r], String] \ Write(r) = 
        try {
            import org.apache.poi.ss.usermodel.Workbook.createCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ r; 
            let Workbook(wb1) = wb;
            createCellStyle(wb1) |> CellStyle |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }

    pub def createFont!(wb: Workbook[r]): Font[r] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createFont(): ##org.apache.poi.ss.usermodel.Font \ r; 
        let Workbook(wb1) = wb;
        createFont(wb1) |> Font

    pub def createDataFormat!(wb: Workbook[r]): DataFormat[r] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createDataFormat(): ##org.apache.poi.ss.usermodel.DataFormat \ r; 
        let Workbook(wb1) = wb;
        createDataFormat(wb1) |> DataFormat

}
