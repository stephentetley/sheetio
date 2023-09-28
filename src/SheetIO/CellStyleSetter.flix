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


mod SheetIO {


    pub type alias CellStyleSetter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.Builder[ka, Internal.Apache.Poi.SS.Usermodel.CellStyle[r], ef, a]



}

mod SheetIO.CellStyleSetter {

    use Graded.Builder.{liftSetter!, liftSetterResult!}

    use SheetIO.CellStyleSetter

    use Internal.Apache.Poi.SS.Usermodel.CellStyle
    use Internal.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Internal.Apache.Poi.SS.Usermodel.Font
    use Internal.Apache.Poi.SS.Usermodel.Font.{Font}


    /// The state param `st` is expected to be mutated...
    /// Blank CellStyles created from the workbook

    pub def runCellStyleSetter!(ma: CellStyleSetter[a, r, ef, a], 
                                st: CellStyle[r]): Result[String, a] \ ef =
        Graded.Builder.runBuilder(ma, st)

    pub def setFont!(font: Font[r]): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setFont!(font))

    
    pub def setWrapText!(wrapped: Bool): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setWrapText!(wrapped))


    pub def setIndention!(indent: Int16): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setIndention!(indent))

    pub def setFillForegroundColor!(color: Int16): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setFillForegroundColor!(color))
  
    pub def setFillBackgroundColor!(color: Int16): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setFillBackgroundColor!(color))


    pub def setShrinkToFit!(shrink: Bool): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setShrinkToFit!(shrink))

    pub def setDataFormat!(value: Int16): CellStyleSetter[ka, r, r, Unit] =
        liftSetter!(Internal.Apache.Poi.SS.Usermodel.CellStyle.setDataFormat!(value))


    // pub def cloneStyleFrom!(x: CellStyle[r]): CellStyleSetter[ka, r, Unit] \ r =
    //     let CellStyle(x1) = x;
    //     liftPrim(r -> SheetIO.Internal.CellStyle.cloneStyleFrom!(r, x1))



}
