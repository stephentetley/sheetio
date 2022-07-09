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


namespace SheetIO/Internal/EvalError {

    use SheetIO/Internal/PrimError.{PrimError};

    pub type alias EvalError = 
        { source :: String
        , exnMessage :: String
        , sheet :: String
        , rowNumber :: Int32
        }

    pub def create(r: Region[r], row: ##org.apache.poi.ss.usermodel.Row, prim: PrimError): EvalError \ Write(r) = 
        use Result.{flatMap};
        let step = () -> {
            let* sheet = SheetIO/Internal/Row.getSheet(r, row);
            let* sheetName = SheetIO/Internal/Sheet.getSheetName(r, sheet);
            let  rowNum = SheetIO/Internal/Row.getRowNum(r, row);
            Ok({sheet = sheetName, rowNumber = rowNum, source = prim.source, exnMessage = prim.exnMessage})
        };
        step() |> Result.getWithDefault({sheet = "***", rowNumber = 0, source = prim.source, exnMessage = prim.exnMessage})


    pub def getText(err: EvalError): String =
        String.unlines(SheetIO/Internal/PrimError.getText(err) :: "Sheet: ${err.sheet}, Row: ${err.rowNumber}" :: Nil)

}

