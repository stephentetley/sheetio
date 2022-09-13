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

namespace Interop/Apache/Poi/Ss/Usermodel {

    /// Mutable / Interface

    pub enum Workbook[_: Region](##org.apache.poi.ss.usermodel.Workbook)

    instance Scoped[Workbook] {
        pub def regionOf(_: Workbook[r]): Region[r] = () as Region[r]
    }
}


namespace Interop/Apache/Poi/Ss/Usermodel/Workbook {

    use Interop/Apache/Poi/Ss/Usermodel.Workbook;
    use Interop/Apache/Poi/Ss/Usermodel.Workbook.{Workbook};
    use Interop/Apache/Poi/Ss/Usermodel.Sheet;
    use Interop/Apache/Poi/Ss/Usermodel.Sheet.{Sheet};

    pub def getHH(wb: Workbook[r]): ##org.apache.poi.ss.usermodel.Workbook =
        let Workbook(wb1) = wb;
        wb1

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

/*
    pub def getSheet(name: String, wb: Workbook[r]): Result[Sheet[r], PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheet(String): ##org.apache.poi.ss.usermodel.Sheet \ r; 
        try {
            let ans = getSheet(wb, name);
            if (Object.isNull(ans))
                Err({source = "getSheet", exnMessage = "Answer is null"})
            else
                Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheet", ex)
        }

    pub def getSheetAt(ix: Int32, wb: Workbook[r]): Result[Sheet[r], PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getSheetAt(Int32): ##org.apache.poi.ss.usermodel.Sheet \ r; 
        try {
            let ans = getSheetAt(wb, ix);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getSheetAt", ex)
        }

    pub def getFontAt(ix: Int32, wb: Workbook[r]): Result[Font[r], PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Workbook.getFontAt(Int32): ##org.apache.poi.ss.usermodel.Font \ r; 
        try {
            let ans = getFontAt(wb, ix);
            if (Object.isNull(ans))
                Err({source = "getFontAt", exnMessage = "Answer is null"})
            else
                Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getFontAt", ex)
        }


    pub def createCellStyle!(wb: Workbook[r]): Result[CellStyle[r], PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createCellStyle(): ##org.apache.poi.ss.usermodel.CellStyle \ r; 
        try {
            let ans = createCellStyle(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createCellStyle", ex)
        }

    pub def createFont!(wb: Workbook[r]): Result[Font[r], PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createFont(): ##org.apache.poi.ss.usermodel.Font \ r; 
        try {
            let ans = createFont(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createFont", ex)
        }

    pub def createDataFormat(wb: Workbook[r]): Result[##org.apache.poi.ss.usermodel.DataFormat, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Workbook.createDataFormat(): ##org.apache.poi.ss.usermodel.DataFormat \ r; 
        try {
            let ans = createDataFormat(wb);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "createDataFormat", ex)
        }
*/
}
