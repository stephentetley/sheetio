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

namespace SheetIO/CursorReader {


    use System.FilePath.{FilePath};

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};


    type alias POIRow = ##org.apache.poi.ss.usermodel.Row

    ///
    /// Returns a new Row Cursor.
    ///
    pub def newRowCursor(src: System.FilePath, sheetName: String): Result[Cursor[POIRow], String] & Impure =
        import new flix.runtime.spt.sheetio.StreamingRowIterator(##java.nio.file.Path, String) as newRowIterator;
        import flix.runtime.spt.sheetio.StreamingRowIterator.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.next() as ffiNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.close() as ffiClose;
        try {
            let FilePath(src1) = src;
            let iter = newRowIterator(src1, sheetName);
            let cursorIter =  {
                  hasNext = () -> ffiHasNext(iter)
                , next = () -> ffiNext(iter)
                , finalize = () -> ffiClose(iter)
            };
            Ok(Cursor.makeCursor(cursorIter))
        } catch {
            case _: ##java.lang.Exception => Err("newRowCursor")
        }

    ///
    /// Returns a new Row Cursor for sheet at `ix`.
    ///
    pub def newRowCursorAt(src: System.FilePath, sheetIndex: Int32): Result[Cursor[POIRow], String] & Impure =
        import new flix.runtime.spt.sheetio.StreamingRowIterator(##java.nio.file.Path, Int32) as newRowIterator;
        import flix.runtime.spt.sheetio.StreamingRowIterator.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.next() as ffiNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.close() as ffiClose;
        try {
            let FilePath(src1) = src;
            let iter = newRowIterator(src1, sheetIndex);
            let cursorIter =  {
                  hasNext = () -> ffiHasNext(iter)
                , next = () -> ffiNext(iter)
                , finalize = () -> ffiClose(iter)
            };
            Ok(Cursor.makeCursor(cursorIter))
        } catch {
            case _: ##java.lang.Exception => Err("newRowCursorAt")
        }

    /// 
    /// At the moment the answer type is not ideal, it would be much nicer to embed the `Result` type 
    /// in the cursor so the answer is `Cursor[a]` or more likely `Cursor[m, a]` where `m` is some
    /// monad.
    /// 
    pub def indexedRowCursor(start: Int32, eval: RowEvaluator[a], cr: Cursor[POIRow]): Cursor[Result[a, String]] = 
        let RowEvaluator(eval1) = eval;
        cr |> Cursor.drop(start) |> Cursor.map(eval1)

    /// 
    /// Likewise this function has a similar not ideal answer type, which is even more cumbersome due to the effect 
    /// (the Cursor type already embeds Impure but here it has to leak out)
    /// 
    /// Because Cursor is implicitly tied to iterating "IO-like" resources (e.g. lines of a file, database results)
    /// it should most likely embed errors itself.
    /// 
    pub def namedRowCursor(headersLine: Int32, eval: NamedRowEvaluator[a], cr: Cursor[POIRow]): Cursor[Result[a, String]] & Impure = 
        let NamedRowEvaluator(eval1) = eval;
        let RowEvaluator(headerEval1) = SheetIO/RowEvaluator.getHeaders();
        let cr1 = if (headersLine > 0) Cursor.drop(headersLine - 1, cr) else cr;
        match Cursor.view(cr1) {
            case Some(hrow, cr2) => {
                let arr = Result.getWithDefault(headerEval1(hrow), Array.new("", 0));
                let lookups = SheetIO/NamedRowEvaluator.makeHeaders(arr); 
                Cursor.map(eval1(lookups), cr2)
            }
            case None => Cursor.empty()
        }

}