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

    /// Mutable

    pub enum Font[_: Region](##org.apache.poi.ss.usermodel.Font)


}

namespace Interop/Apache/Poi/SS/Usermodel/Font {

    use Interop/Apache/Poi/SS/Usermodel.Font;
    use Interop/Apache/Poi/SS/Usermodel.Font.{Font};

    pub def getHH(font: Font[r]): ##org.apache.poi.ss.usermodel.Font =
        let Font(font1) = font;
        font1

    pub def wrapHH(_: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Font[r] =
        Font(font)

    pub def setFontName!(name: String, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontName(String): Unit \ r;
        let Font(font1) = font;
        setFontName(font1, name)


    pub def setFontHeight!(height: Int16, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontHeight(Int16): Unit \ r;
        let Font(font1) = font;
        setFontHeight(font1, height)

    pub def setFontHeightInPoints!(height: Int16, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontHeightInPoints(Int16): Unit \ r; 
        let Font(font1) = font;
        setFontHeightInPoints(font1, height)

    
    pub def getBold(font: Font[r]): Bool \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getBold(): Bool \ r;
        let Font(font1) = font;
        getBold(font1)

    pub def setBold!(bold: Bool, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setBold(Bool): Unit \ r;
        let Font(font1) = font;
        setBold(font1, bold)

    pub def setItalic!(italic: Bool, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setItalic(Bool): Unit \ r; 
        let Font(font1) = font;
        setItalic(font1, italic)

    
    pub def getItalic(font: Font[r]): Bool \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getItalic(): Bool \ r; 
        let Font(font1) = font;
        getItalic(font1)

    pub def setStrikeout!(strikeout: Bool, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setStrikeout(Bool): Unit \ r;
        let Font(font1) = font;
        setStrikeout(font1, strikeout)
    
    pub def getStrikeout(font: Font[r]): Bool \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getStrikeout(): Bool \ r; 
        let Font(font1) = font;
        getStrikeout(font1)

    pub def setColor!(color: Int16, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setColor(Int16): Unit \ r;
        let Font(font1) = font;
        setColor(font1, color)

    pub def setTypeOffset!(offset: Int16, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setTypeOffset(Int16): Unit \ r; 
        let Font(font1) = font;
        setTypeOffset(font1, offset)

    pub def setUnderline!(underline: Int8, font: Font[r]): Unit \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setUnderline(Int8): Unit \ r; 
        let Font(font1) = font;
        setUnderline(font1, underline)

}
