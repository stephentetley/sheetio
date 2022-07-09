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

    use SheetIO/Internal/PrimError.{PrimError};
    
    pub def setFontName!(r: Region[r],
                        value: String, 
                        font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontName(String): Unit & r; 
        try {
            let ans = setFontName(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setFontName!", ex)
        }


    pub def setFontHeight!(r: Region[r],
                            value: Int16,
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontHeight(Int16): Unit & r; 
        try {
            let ans = setFontHeight(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setFontHeight!", ex)
        }

    pub def setFontHeightInPoints!(r: Region[r],
                                    value: Int16,
                                    font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setFontHeightInPoints(Int16): Unit & r; 
        try {
            let ans = setFontHeightInPoints(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setFontHeightInPoints!", ex)
        }

    
    pub def getBold(r: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getBold(): Bool & r;
        try {
            let ans = getBold(font);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getBold", ex)
        }

    pub def setBold!(r: Region[r], 
                    value: Bool, 
                    font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setBold(Bool): Unit & r; 
        try {
            let ans = setBold(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setBold!", ex)
        }

    pub def setItalic!(r: Region[r],
                        value: Bool,
                        font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setItalic(Bool): Unit & r; 
        try {
            let ans = setItalic(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setItalic!", ex)
        }

    
    pub def getItalic(r: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getItalic(): Bool & r; 
        try {
            getItalic(font) |> Ok
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getItalic", ex)
        }

    pub def setStrikeout!(r: Region[r],
                            value: Bool, 
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setStrikeout(Bool): Unit & r; 
        try {
            let ans = setStrikeout(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setStrikeout!", ex)
        }
    
    pub def getStrikeout(r: Region[r], font: ##org.apache.poi.ss.usermodel.Font): Result[Bool, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Font.getStrikeout(): Bool & r; 
        try {
            let ans = getStrikeout(font);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getStrikeout", ex)
        }

    pub def setColor!(r: Region[r],
                        value: Int16,
                        font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setColor(Int16): Unit & r; 
        try {
            let ans = setColor(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setColor!", ex)
        }

    pub def setTypeOffset!(r: Region[r],
                            value: Int16,
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setTypeOffset(Int16): Unit & r; 
        try {
            let ans = setTypeOffset(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setTypeOffset!", ex)
        }

    pub def setUnderline!(r: Region[r],
                            value: Int8, 
                            font: ##org.apache.poi.ss.usermodel.Font): Result[Unit, PrimError] \ Write(r) = 
        import org.apache.poi.ss.usermodel.Font.setUnderline(Int8): Unit & r; 
        try {
            let ans = setUnderline(font, value);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "setUnderline!", ex)
        }

}
