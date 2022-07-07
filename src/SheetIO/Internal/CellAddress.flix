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

namespace SheetIO/Internal/CellAddress {


    pub def new(_: Region[r], s: String): ##org.apache.poi.ss.util.CellAddress \ Write(r) = 
        import new org.apache.poi.ss.util.CellAddress(String): ##org.apache.poi.ss.util.CellAddress & r as newCellAddress; 
        newCellAddress(s)


    pub def getColumn(_: Region[r],
                        addr: ##org.apache.poi.ss.util.CellAddress): Result[Int32, String] \Read(r) = 
        import org.apache.poi.ss.util.CellAddress.getColumn(): Int32 & r; 
        try {
            let ans = getColumn(addr);
            Ok(ans)
        } catch {
            case ex: ##java.lang.Exception => 
                import java.lang.Throwable.getMessage(): String & r;                
                Err("getColumn: ${getMessage(ex)}") 
        }

}

