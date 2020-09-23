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

namespace SheetReader/Internal/Sheet {

    pub def getFirstRowNum(ws: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getFirstRowNum(); 
        try {
            let ans = ws.getFirstRowNum();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getFirstRowNum")) 
        }

    pub def getLastRowNum(ws: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getLastRowNum(); 
        try {
            let ans = ws.getLastRowNum();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getLastRowNum")) 
        }

    pub def getPhysicalNumberOfRows(ws: ##org.apache.poi.ss.usermodel.Sheet): Result[Int32, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getPhysicalNumberOfRows(); 
        try {
            let ans = ws.getPhysicalNumberOfRows();
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getPhysicalNumberOfRows")) 
        }


    pub def getRow(ws: ##org.apache.poi.ss.usermodel.Sheet,
                   ix: Int32): Result[##org.apache.poi.ss.usermodel.Row, System.Error] & Impure = 
        import org.apache.poi.ss.usermodel.Sheet.getRow(Int32); 
        try {
            let ans = ws.getRow(ix);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err(Generic("getRow")) 
        }

}