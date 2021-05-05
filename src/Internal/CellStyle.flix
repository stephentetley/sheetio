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

namespace SheetIO/Internal/CellStyle {

    pub def setFont(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                    font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setFont(##org.apache.poi.ss.usermodel.Font); 
        try {
            let ans = setFont(style, font);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setFont")
        }

    pub def getFontIndexAsInt(style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Int32, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.getFontIndexAsInt(); 
        try {
            let ans = getFontIndexAsInt(style);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("getFontIndexAsInt")
        }

    pub def setWrapText(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                        value: Bool): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setWrapText(Bool); 
        try {
            let ans = setWrapText(style, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setWrapText")
        }
    
    pub def setIndention(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                         value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setIndention(Int16); 
        try {
            let ans = setIndention(style, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setIndention")
        }
    
    
    pub def setFillForegroundColor(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                                   value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillForegroundColor(Int16); 
        try {
            let ans = setFillForegroundColor(style, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setFillForegroundColor")
        }

    /// Note must set foreground color first
    pub def setFillBackgroundColor(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                                   value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillBackgroundColor(Int16); 
        try {
            let ans = setFillBackgroundColor(style, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setFillBackgroundColor")
        }


    pub def setShrinkToFit(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                           value: Bool): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setShrinkToFit(Bool); 
        try {
            let ans = setShrinkToFit(style, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setShrinkToFit") 
        }

    pub def setDataFormat(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                          value: Int16): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setDataFormat(Int16); 
        try {
            let ans = setDataFormat(style, value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("setDataFormat") 
        }

    pub def cloneStyleFrom(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                           styleToClone: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.cloneStyleFrom(##org.apache.poi.ss.usermodel.CellStyle); 
        try {
            let ans = cloneStyleFrom(style, styleToClone);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("cloneStyleFrom") 
        }

}
