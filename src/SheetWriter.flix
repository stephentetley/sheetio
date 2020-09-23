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

use SheetIO.SheetWriter;
use SheetIO.SheetWriter.{SheetWriter};


namespace SheetIO {

    pub opaque type SheetWriter[a] = ##org.apache.poi.ss.usermodel.Workbook ~> Result[a, System.Error]

}


namespace SheetIO/SheetWriter {

    pub def runXlsx(ma: SheetWriter[a], path: System.FilePath): Result[a, System.Error] & Impure = 
        use SheetIO/Internal/Prim.{newXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let outpath = System/FilePath.toString(path);
        let* wb     = newXSSFWorkbook();
        let* ans    = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook);
        let* _      = writeXSSFWorkbook(wb, outpath);
        let* _      = closeXSSFWorkbook(wb);
        Ok(ans)

    def apply1(ma: SheetWriter[a], wb: ##org.apache.poi.ss.usermodel.Workbook): Result[a, System.Error] & Impure = 
        let SheetWriter(x1) = ma;
        x1(wb)

    pub def return(x: a): SheetWriter[a] = 
        SheetWriter(_ -> Ok(x) as & Impure)

}