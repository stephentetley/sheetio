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


namespace SheetIO/Internal/Font {

    pub def setFontName(font: ##org.apache.poi.ss.usermodel.Font,
                        value: String): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setFontName(String); 
        try {
            let ans = font.setFontName(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setFontName")
        }

    pub def setFontHeight(font: ##org.apache.poi.ss.usermodel.Font,
                          value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setFontHeight(Int16); 
        try {
            let ans = font.setFontHeight(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setFontHeight")
        }

    pub def setFontHeightInPoints(font: ##org.apache.poi.ss.usermodel.Font,
                                  value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setFontHeightInPoints(Int16); 
        try {
            let ans = font.setFontHeightInPoints(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setFontHeightInPoints")
        }

    
    pub def getBold(font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.getBold(); 
        try {
            let ans = getBold(font);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getBold")
        }

    pub def setBold(font: ##org.apache.poi.ss.usermodel.Font,
                    value: Bool): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setBold(Bool); 
        try {
            let ans = font.setBold(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setBold")
        }

    pub def setItalic(font: ##org.apache.poi.ss.usermodel.Font,
                      value: Bool): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setItalic(Bool); 
        try {
            let ans = font.setItalic(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setItalic")
        }

    
    pub def getItalic(font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.getItalic(); 
        try {
            let ans = getItalic(font);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getItalic")
        }

    pub def setStrikeout(font: ##org.apache.poi.ss.usermodel.Font,
                         value: Bool): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setStrikeout(Bool); 
        try {
            let ans = font.setStrikeout(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setStrikeout")
        }
    
    pub def getStrikeout(font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.getStrikeout(); 
        try {
            let ans = getStrikeout(font);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getStrikeout")
        }

    pub def setColor(font: ##org.apache.poi.ss.usermodel.Font,
                     value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setColor(Int16); 
        try {
            let ans = font.setColor(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setColor")
        }

    pub def setTypeOffset(font: ##org.apache.poi.ss.usermodel.Font,
                          value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setTypeOffset(Int16); 
        try {
            let ans = font.setTypeOffset(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setTypeOffset")
        }

    pub def setUnderline(font: ##org.apache.poi.ss.usermodel.Font,
                         value: Int8): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setUnderline(Int8); 
        try {
            let ans = font.setUnderline(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setUnderline")
        }

}