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

namespace SheetIO {


    pub type alias EvalError = 
        { message :: String
        , sheet :: String
        , rowNumber :: Int32
        }

}

namespace SheetIO/EvalError {

    use SheetIO.EvalError;

    pub def create(r: Region[r], row: ##org.apache.poi.ss.usermodel.Row, msg: String): EvalError \ Read(r) = 
        let rownum = SheetIO/Internal/Row.getRowNum(r, row);
        match Result.flatMap(SheetIO/Internal/Sheet.getSheetName(r), SheetIO/Internal/Row.getSheet(r, row)) {
            case Err(_) => {message = msg, sheet = "***", rowNumber = rownum}
            case Ok(s1) => {message = msg, sheet = s1, rowNumber = rownum}
        }

    pub def toString(err: EvalError): String = 
        String.unlines(err.message :: "Sheet: ${err.sheet}, Row: ${err.rowNumber}" :: Nil)

}

