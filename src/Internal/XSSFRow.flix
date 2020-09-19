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

namespace SheetReader/Internal/XSSFRow {

    pub def getRowNum(row: ##org.apache.poi.xssf.usermodel.XSSFRow): Int32 & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFRow.getRowNum(); 
        row.getRowNum()
    
    pub def getCell(row: ##org.apache.poi.xssf.usermodel.XSSFRow, 
                    ix: Int32): Result[##org.apache.poi.xssf.usermodel.XSSFCell, System.Error] & Impure = 
        import org.apache.poi.xssf.usermodel.XSSFRow.getCell(Int32); 
        try {
            let ans = row.getCell(ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getCell")) 
        }

}