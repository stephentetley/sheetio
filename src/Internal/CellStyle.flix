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

use System.Error.{Generic};

namespace SheetIO/Internal/CellStyle {

    pub def setFont(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                    font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setFont(##org.apache.poi.ss.usermodel.Font); 
        try {
            let ans = style.setFont(font);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setFont")) 
        }

    pub def setWrapText(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                        value: Bool): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setWrapText(Bool); 
        try {
            let ans = style.setWrapText(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setWrapText")) 
        }
    
    pub def setIndention(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                         value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setIndention(Int16); 
        try {
            let ans = style.setIndention(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setIndention")) 
        }
    
    
    pub def setFillForegroundColor(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                                   value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillForegroundColor(Int16); 
        try {
            let ans = style.setFillForegroundColor(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setFillForegroundColor")) 
        }

    /// Note must set foreground color first
    pub def setFillBackgroundColor(style: ##org.apache.poi.ss.usermodel.CellStyle, 
                                   value: Int16): Result[Unit, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillBackgroundColor(Int16); 
        try {
            let ans = style.setFillBackgroundColor(value);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("setFillBackgroundColor")) 
        }
}