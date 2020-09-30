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

    use System.Error.{Generic};

    pub def setFontName(font: ##org.apache.poi.ss.usermodel.Font,
                        value: String): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setFontName(String); 
        try {
            let ans = font.setFontName(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setFontName")) 
        }

    pub def setFontHeight(font: ##org.apache.poi.ss.usermodel.Font,
                          value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setFontHeight(Int16); 
        try {
            let ans = font.setFontHeight(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setFontHeight")) 
        }

    pub def setFontHeightInPoints(font: ##org.apache.poi.ss.usermodel.Font,
                                  value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setFontHeightInPoints(Int16); 
        try {
            let ans = font.setFontHeightInPoints(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setFontHeightInPoints")) 
        }


    pub def setBold(font: ##org.apache.poi.ss.usermodel.Font,
                    value: Bool): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setBold(Bool); 
        try {
            let ans = font.setBold(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setBold")) 
        }

    pub def setItalic(font: ##org.apache.poi.ss.usermodel.Font,
                      value: Bool): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setItalic(Bool); 
        try {
            let ans = font.setItalic(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setItalic")) 
        }


    pub def setStrikeout(font: ##org.apache.poi.ss.usermodel.Font,
                         value: Bool): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setStrikeout(Bool); 
        try {
            let ans = font.setStrikeout(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setStrikeout")) 
        }

    pub def setColor(font: ##org.apache.poi.ss.usermodel.Font,
                     value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setColor(Int16); 
        try {
            let ans = font.setColor(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setColor")) 
        }

    pub def setTypeOffset(font: ##org.apache.poi.ss.usermodel.Font,
                          value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setTypeOffset(Int16); 
        try {
            let ans = font.setTypeOffset(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setTypeOffset")) 
        }

    pub def setUnderline(font: ##org.apache.poi.ss.usermodel.Font,
                         value: Int8): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Font.setUnderline(Int8); 
        try {
            let ans = font.setUnderline(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setUnderline")) 
        }

}