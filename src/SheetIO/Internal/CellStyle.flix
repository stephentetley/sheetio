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

    pub def setFont!(_: Region[r],
                    font: ##org.apache.poi.ss.usermodel.Font,
                    style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] & r = 
        import org.apache.poi.ss.usermodel.CellStyle.setFont(##org.apache.poi.ss.usermodel.Font): Unit & r; 
        try {
            let ans = setFont(style, font);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setFont: ${getMessage(ex)}")
        }

    pub def getFontIndexAsInt(_: Region[r], 
                                style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Int32, String] & r = 
        import org.apache.poi.ss.usermodel.CellStyle.getFontIndexAsInt(): Int32 & r; 
        try {
            let ans = getFontIndexAsInt(style);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getFontIndexAsInt: ${getMessage(ex)}")            
        }

    pub def setWrapText!(_: Region[r], 
                        value: Bool,
                        style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.setWrapText(Bool): Unit & r; 
        try {
            let ans = setWrapText(style, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setWrapText: ${getMessage(ex)}")              
        }
    
    pub def setIndention!(_: Region[r], 
                            value: Int16,
                            style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.setIndention(Int16): Unit & r; 
        try {
            let ans = setIndention(style, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setIndention: ${getMessage(ex)}")              
        }
    
    
    pub def setFillForegroundColor!(_: Region[r], 
                                    value: Int16, 
                                    style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillForegroundColor(Int16): Unit & r; 
        try {
            let ans = setFillForegroundColor(style, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setFillForegroundColor: ${getMessage(ex)}")              
        }

    /// Note must set foreground color first
    pub def setFillBackgroundColor!(_: Region[r], 
                                   value: Int16,
                                   style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.setFillBackgroundColor(Int16): Unit & r; 
        try {
            let ans = setFillBackgroundColor(style, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setFillBackgroundColor: ${getMessage(ex)}")              
        }


    pub def setShrinkToFit!(_: Region[r], 
                           value: Bool, 
                           style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.setShrinkToFit(Bool): Unit & r; 
        try {
            let ans = setShrinkToFit(style, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setShrinkToFit: ${getMessage(ex)}")  
        }

    pub def setDataFormat!(_: Region[r], 
                            value: Int16, 
                            style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.setDataFormat(Int16): Unit & r; 
        try {
            let ans = setDataFormat(style, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setDataFormat: ${getMessage(ex)}")  
        }

    pub def cloneStyleFrom!(_: Region[r], 
                            styleToClone: ##org.apache.poi.ss.usermodel.CellStyle, 
                            style: ##org.apache.poi.ss.usermodel.CellStyle): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.CellStyle.cloneStyleFrom(##org.apache.poi.ss.usermodel.CellStyle): Unit & r; 
        try {
            let ans = cloneStyleFrom(style, styleToClone);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("cloneStyleFrom: ${getMessage(ex)}")              
        }

}
