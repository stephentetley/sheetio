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
    use SheetIO.SheetWriter;
    
    /// Note - this isn't very satisfactory. Ideally WorkSheetWriter would not be 
    /// parametric on row so we could put them in a list.

    /// TODO - rows should be t[a] with Foldable[t]...
    type alias WorkSheetWriter[row] = 
        { sheetName: String
        , headers: Option[Array[String]]
        , setter: row -> RowSetter[Unit]
        , rows: List[row]
        } 

    /// Write a workbook containing a single sheet
    pub def writeWorkSheet(sheetWriter: WorkSheetWriter[a], path: System.FilePath): Result[Unit, String] & Impure =
        SheetIO/SheetWriter.runXlsx(writeSheet(sheetWriter), path)

    /// It is not viable to write a list of WorkSheetWriter as a list of WorkSheetWriter's are unlikely 
    /// to have the same row type.
    /// To write multiple sheets we need to be able to append sheets...

    /// Add a Sheet to an existing Workbook
    pub def addWorkSheet(sheetWriter: WorkSheetWriter[a], path: System.FilePath): Result[Unit, String] & Impure =
        SheetIO/SheetWriter.runAppendXlsx(writeSheet(sheetWriter), path)


    def writeSheet(sheetWriter: WorkSheetWriter[a]): SheetWriter[Unit] = 
        use SheetIO/SheetWriter.{flatMap, return, createSheet, allRows, allRowsWithHeaders};
        match sheetWriter.headers {
            case None => {
                let* s1     = createSheet(sheetWriter.sheetName);
                let* _      = allRows(s1, sheetWriter.rows, sheetWriter.setter);
                return()
            }
            case Some(headers) => {
                let* s1     = createSheet(sheetWriter.sheetName);
                let* _      = allRowsWithHeaders(s1, headers, sheetWriter.rows, sheetWriter.setter);
                return()
            }
        }


}
