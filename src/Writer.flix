/*
 * Copyright 2021 Stephen Tetley
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



namespace SheetIO/Writer {

    use SheetIO.RowSetter;
    
    

    pub def writeAllRows(path: System.FilePath, setter: a -> RowSetter[Unit], rows: List[a]): Result[Unit, String] & Impure =
        use SheetIO/SheetWriter.{flatMap, return, createSheet, allRows};
        let writeStep = () -> { 
            let* s1     = createSheet("Sheet1");
            let* _      = allRows(s1, rows, setter);
            return()
        };
        SheetIO/SheetWriter.runXlsx(writeStep(), path)


    pub def writeAllRowsWithHeaders(path: System.FilePath, setter: a-> RowSetter[Unit], headers: Array[String], rows: List[a]): Result[Unit, String] & Impure =
        use SheetIO/SheetWriter.{flatMap, return, createSheet, allRowsWithHeaders};
        let writeStep = () -> { 
            let* s1     = createSheet("Sheet1");
            let* _      = allRowsWithHeaders(s1, headers, rows, setter);
            return()
        };
        SheetIO/SheetWriter.runXlsx(writeStep(), path)

}
