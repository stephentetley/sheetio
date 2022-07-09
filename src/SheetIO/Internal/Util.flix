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

namespace SheetIO/Internal/Util {

    use SheetIO/Internal/PrimError.{PrimError};


    /// Caller should test we have a NUMERIC cell first...
    pub def isCellDateFormatted(r: Region[r], 
                                cell: ##org.apache.poi.ss.usermodel.Cell): Result[Bool, PrimError] \ Read(r) = 
        import static org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(##org.apache.poi.ss.usermodel.Cell): Bool & r;
        try { 
            let ans = isCellDateFormatted(cell);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "isCellDateFormatted", ex)
        }

    /// Orphan - DataFormat not used enough to warrant a module.
    pub def getFormat(r: Region[r],
                        format: String, 
                        dfmt: ##org.apache.poi.ss.usermodel.DataFormat): Result[Int16, PrimError] \ Read(r) = 
        import org.apache.poi.ss.usermodel.DataFormat.getFormat(String): Int16 & r; 
        try {
            let ans = getFormat(dfmt, format);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => SheetIO/Internal/PrimError.primErr(r, "getFormat", ex)
        }

}

