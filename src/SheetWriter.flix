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

use SheetIO.CellType;
use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};


namespace SheetIO {

    pub opaque type SheetWriter[a] = ##org.apache.poi.ss.usermodel.Workbook ~> Result[a, System.Error]

}


namespace SheetReader/SheetWriter {

    // pub def runXlsx(ma: WbMonad[a], path: System.FilePath): Result[a, System.Error] & Impure = 
    //     use SheetReader/Internal/Prim.{newXSSFWorkbook, closeXSSFWorkbook};
    //     use Result.{flatMap};
    //     let path1 = System/FilePath.toString(path);
    //     match newXSSFWorkbook(path1) { 
    //         case Err(e) => Err(e)
    //         case Ok(wb) => {
    //             let ans = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook);
    //             let* _ = closeXSSFWorkbook(wb);
    //             ans         
    //         }
    //     }

    
}