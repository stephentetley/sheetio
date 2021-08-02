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

/// Streaming reader for large lfiles with homogeneous rows.

namespace SheetIO/CursorReader {

    use X.Cursor;
    use System.FilePath.{FilePath};

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};


    type alias POIRow = ##org.apache.poi.ss.usermodel.Row

    ///
    /// Returns a new Row Cursor.
    ///
    pub def newRowCursor(src: System.FilePath, sheetName: String): Cursor[String, POIRow] & Impure =
        import flix.runtime.spt.sheetio.StreamingRowCursor:createCursorWithSheetName(##java.nio.file.Path, String) as newRowCursor;
        import flix.runtime.spt.sheetio.StreamingRowCursor.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.StreamingRowCursor.next() as ffiNext;
        import flix.runtime.spt.sheetio.StreamingRowCursor.close() as ffiClose;
        try {
            let FilePath(src1) = src;
            let iter = newRowCursor(src1, sheetName);
            let cursorAlg =  {
                  hasNext   = () -> ffiHasNext(iter)
                , next      = X/Cursor.guardThunk(_ -> ffiNext(iter), "next failed")
                , finalize  = () -> ffiClose(iter)
            };
            X/Cursor.makeCursor(cursorAlg)
        } catch {
            case _: ##java.lang.Exception => X/Cursor.failure("newRowCursor")
        }

    ///
    /// Returns a new Row Cursor for sheet at `ix`.
    ///
    pub def newRowCursorAt(src: System.FilePath, sheetIndex: Int32): Cursor[String, POIRow] & Impure =
        import flix.runtime.spt.sheetio.StreamingRowCursor:createCursorWithSheetNumber(##java.nio.file.Path, Int32) as newRowCursor;
        import flix.runtime.spt.sheetio.StreamingRowCursor.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.StreamingRowCursor.next() as ffiNext;
        import flix.runtime.spt.sheetio.StreamingRowCursor.close() as ffiClose;
        try {
            let FilePath(src1) = src;
            let iter = newRowCursor(src1, sheetIndex);
            let cursorAlg =  {
                  hasNext   = () -> ffiHasNext(iter)
                , next      = X/Cursor.guardThunk(_ -> ffiNext(iter), "next failed")
                , finalize  = () -> ffiClose(iter)
            };
            X/Cursor.makeCursor(cursorAlg)
        } catch {
            case _: ##java.lang.Exception => X/Cursor.failure("newRowCursorAt")
        }

    /// 
    /// Improved answer type...
    /// 
    pub def indexedRowCursor(start: Int32, eval: RowEvaluator[a], src: Cursor[String, POIRow]): Cursor[String, a] & Impure = 
        let RowEvaluator(eval1) = eval;
        let fmt = SheetIO/Internal/DataFormatter.new();
        src |> X/Cursor.drop(start) |> X/Cursor.mapResult(eval1(fmt))

    /// 
    /// Because Cursor is implicitly tied to iterating "IO-like" resources (e.g. lines of a file, database results)
    /// it embeds errors.
    /// 
    pub def namedRowCursor(headersLine: Int32, eval: NamedRowEvaluator[a], src: Cursor[String, POIRow]): Cursor[String, a] & Impure = 
        let NamedRowEvaluator(eval1) = eval;
        let RowEvaluator(headerEval1) = SheetIO/RowEvaluator.getHeaders();
        let fmt = SheetIO/Internal/DataFormatter.new();
        let src1 = if (headersLine > 0) X/Cursor.drop(headersLine - 1, src) else src;
        match X/Cursor.head!(src1) {
            case Some(hrow) => {
                let arr = Result.getWithDefault(headerEval1(fmt, hrow), Array.new("", 0));
                let lookups = SheetIO/NamedRowEvaluator.makeHeaders(arr); 
                X/Cursor.mapResult(eval1(lookups), src1)
            }
            case None => { 
                X/Cursor.finalize!(src1);
                X/Cursor.failure("Cannot find headers")
            }
        }

}
