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

    pub def setFontName!(_: Region[r],
                        value: String, 
                        font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontName(String): Unit & r; 
        try {
            let ans = setFontName(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setFontName!: ${getMessage(ex)}")             
        }


    pub def setFontHeight!(_: Region[r],
                            value: Int16,
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontHeight(Int16): Unit & r; 
        try {
            let ans = setFontHeight(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setFontHeight!: ${getMessage(ex)}")            
        }

    pub def setFontHeightInPoints!(_: Region[r],
                                    value: Int16,
                                    font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontHeightInPoints(Int16): Unit & r; 
        try {
            let ans = setFontHeightInPoints(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setFontHeightInPoints!: ${getMessage(ex)}")
        }

    
    pub def getBold(_: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getBold(): Bool & r;
        try {
            let ans = getBold(font);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getBold: ${getMessage(ex)}")
        }

    pub def setBold!(_: Region[r], 
                    value: Bool, 
                    font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setBold(Bool): Unit & r; 
        try {
            let ans = setBold(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setBold!: ${getMessage(ex)}")
        }

    pub def setItalic!(_: Region[r],
                        value: Bool,
                        font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setItalic(Bool): Unit & r; 
        try {
            let ans = setItalic(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setItalic!: ${getMessage(ex)}")
        }

    
    pub def getItalic(_: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getItalic(): Bool & r; 
        try {
            getItalic(font) |> Ok
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getItalic: ${getMessage(ex)}")
        }

    pub def setStrikeout!(_: Region[r],
                            value: Bool, 
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setStrikeout(Bool): Unit & r; 
        try {
            let ans = setStrikeout(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setStrikeout!: ${getMessage(ex)}")
        }
    
    pub def getStrikeout(_: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, String] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getStrikeout(): Bool & r; 
        try {
            let ans = getStrikeout(font);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getStrikeout: ${getMessage(ex)}")
        }

    pub def setColor!(_: Region[r],
                        value: Int16,
                        font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setColor(Int16): Unit & r; 
        try {
            let ans = setColor(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setColor!: ${getMessage(ex)}")
        }

    pub def setTypeOffset!(_: Region[r],
                            value: Int16,
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setTypeOffset(Int16): Unit & r; 
        try {
            let ans = setTypeOffset(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setTypeOffset!: ${getMessage(ex)}")
        }

    pub def setUnderline!(_: Region[r],
                            value: Int8, 
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, String] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setUnderline(Int8): Unit & r; 
        try {
            let ans = setUnderline(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("setUnderline!: ${getMessage(ex)}")
        }

}
