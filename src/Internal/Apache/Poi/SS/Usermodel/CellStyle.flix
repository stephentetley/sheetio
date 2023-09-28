/*
 * Copyright 2020 Stephen Tetley
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

    /// Mutable
    /// Interface
    pub enum CellStyle[_: Region](##org.apache.poi.ss.usermodel.CellStyle)

    

}

mod Internal.Apache.Poi.SS.Usermodel.CellStyle {

    use Internal.Apache.Poi.SS.Usermodel.CellStyle
    use Internal.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Internal.Apache.Poi.SS.Usermodel.Font
    use Internal.Apache.Poi.SS.Usermodel.Font.{Font}

    pub def getHH(cs: CellStyle[r]): ##org.apache.poi.ss.usermodel.CellStyle =
        let CellStyle(cs1) = cs;
        cs1

   
    pub def wrapHH(_: Region[r], cs: ##org.apache.poi.ss.usermodel.CellStyle): CellStyle[r] =
        CellStyle(cs)

    /// Blank / fresh CellStyles created from the workbook

    pub def setFont!(font: Font[r], cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setFont(##org.apache.poi.ss.usermodel.Font): Unit \ r; 
        let Font(font1) = font;
        let CellStyle(cs1) = cs;
        setFont(cs1, font1)


    pub def getFontIndex(cs: CellStyle[r]): Int32 \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.getFontIndex(): Int32 \ r; 
        let CellStyle(cs1) = cs;
        getFontIndex(cs1)
        

    pub def setWrapText!(wrapped: Bool, cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setWrapText(Bool): Unit \ r; 
        let CellStyle(cs1) = cs;
        setWrapText(cs1, wrapped)

 
    pub def setIndention!(indent: Int16, cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setIndention(Int16): Unit \ r; 
        let CellStyle(cs1) = cs;
        setIndention(cs1, indent)

    

    pub def setFillForegroundColor!(color: Int16, cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillForegroundColor(Int16): Unit \ r;
        let CellStyle(cs1) = cs;
        setFillForegroundColor(cs1, color)

    /// Note must set foreground color first
    pub def setFillBackgroundColor!(color: Int16, cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillBackgroundColor(Int16): Unit \ r; 
        let CellStyle(cs1) = cs;
        setFillBackgroundColor(cs1, color)


    pub def setShrinkToFit!(shrink: Bool, cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setShrinkToFit(Bool): Unit \ r; 
        let CellStyle(cs1) = cs;
        setShrinkToFit(cs1, shrink)

    pub def setDataFormat!(value: Int16, cs: CellStyle[r]): Unit \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.setDataFormat(Int16): Unit \ r; 
        let CellStyle(cs1) = cs;
        setDataFormat(cs1, value)

/*
    pub def cloneStyleFrom!(r: Region[r], 
                            styleToClone: ##org.apache.poi.ss.usermodel.CellStyle, 
                            cs: CellStyle[r]): Result[Unit, PrimError] \ r = 
        import org.apache.poi.ss.usermodel.CellStyle.cloneStyleFrom(##org.apache.poi.ss.usermodel.CellStyle): Unit \ r; 
        try {
            let ans = cloneStyleFrom(style, styleToClone);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO.Internal.PrimError.primErr(r, "cloneStyleFrom!", ex)
        }
*/
}
