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

/// Streaming reader for large files with homogeneous rows.

namespace SheetIO/StreamingReader {

    use X.Cursor;
    use System.FilePath.{FilePath};

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};


    type alias POIRow = ##org.apache.poi.ss.usermodel.Row


    pub opaque type StreamingReaderHandle[a] = (##flix.runtime.spt.sheetio.StreamingRowCursor, Iterator[Result[a, String]])

    pub def close(handle: StreamingReaderHandle[a]): Result[Unit, String] & Impure =
        import flix.runtime.spt.sheetio.StreamingRowCursor.close();
        let StreamingReaderHandle((x1, _)) = handle;
        try {
            close(x1);
            Ok()
        } catch {
            case _: ##java.lang.Exception => Err("close")
        }

    pub def getIterator(handle: StreamingReaderHandle[a]): Iterator[Result[a, String]] =
        let StreamingReaderHandle((_, i1)) = handle;
        i1


    ///
    /// Returns a new StreamingReader Handle (iterator and a "close" handle).
    ///
    pub def newStreamingReaderHandle(src: System.FilePath, sheetName: String): Result[StreamingReaderHandle[POIRow], String] & Impure =
        import flix.runtime.spt.sheetio.StreamingRowCursor:createCursorWithSheetName(##java.nio.file.Path, String);
        import flix.runtime.spt.sheetio.StreamingRowCursor.hasNext();
        import flix.runtime.spt.sheetio.StreamingRowCursor.next();
        try {
            let FilePath(src1) = src;
            let primIter = createCursorWithSheetName(src1, sheetName);
            let iterDone = () -> if (hasNext(primIter)) false else true;
            let iterNext = () -> {
                try {
                    Ok(next(primIter))
                } catch {
                    case _: ##java.lang.Exception => Err("StreamingReader[next]")
                }
            };
            let iter1 = Iterator(iterDone, iterNext);
            let hdl = StreamingReaderHandle((primIter, iter1));
            Ok(hdl)
        } catch {
            case _: ##java.lang.Exception => Err("newStreamingReaderHandle")
        }
 

    ///
    /// Returns a new StreamingReader Handle for sheet at `ix`.
    ///
    pub def newStreamingReaderHandleWithIndex(src: System.FilePath, sheetIndex: Int32): Result[StreamingReaderHandle[POIRow], String] & Impure =
        import flix.runtime.spt.sheetio.StreamingRowCursor:createCursorWithSheetNumber(##java.nio.file.Path, Int32);
        import flix.runtime.spt.sheetio.StreamingRowCursor.hasNext();
        import flix.runtime.spt.sheetio.StreamingRowCursor.next();
        try {
            let FilePath(src1) = src;
            let primIter = createCursorWithSheetNumber(src1, sheetIndex);
            let iterDone = () -> if (hasNext(primIter)) false else true;
            let iterNext = () -> {
                try {
                    Ok(next(primIter))
                } catch {
                    case _: ##java.lang.Exception => Err("StreamingReader[next]")
                }
            };
            let iter1 = Iterator(iterDone, iterNext);
            let hdl = StreamingReaderHandle((primIter, iter1));
            Ok(hdl)
        } catch {
            case _: ##java.lang.Exception => Err("newStreamingReaderHandleWithIndex")
        }


    /// 
    /// Improved answer type...
    /// 
    pub def indexedRowIterator(start: Int32, eval: RowEvaluator[a], src: Iterator[Result[POIRow, String]]): Iterator[Result[a, String]] & Impure = 
        let RowEvaluator(eval1) = eval;
        let fmt = SheetIO/Internal/DataFormatter.new();
        src |> Iterator.drop(start) |> Iterator.map(Result.flatMap(eval1(fmt)))

    /// 
    /// Because Cursor is implicitly tied to iterating "IO-like" resources (e.g. lines of a file, database results)
    /// it embeds errors.
    /// 
    /// pub def namedRowCursor(headersLine: Int32, eval: NamedRowEvaluator[a], src: Cursor[String, POIRow]): Cursor[String, a] & Impure = 
    ///     let NamedRowEvaluator(eval1) = eval;
    ///     let RowEvaluator(headerEval1) = SheetIO/RowEvaluator.getHeaders();
    ///     let fmt = SheetIO/Internal/DataFormatter.new();
    ///     let src1 = if (headersLine > 0) X/Cursor.drop(headersLine - 1, src) else src;
    ///     match X/Cursor.head!(src1) {
    ///         case Some(hrow) => {
    ///             let arr = Result.getWithDefault(headerEval1(fmt, hrow), Array.new("", 0));
    ///             let lookups = SheetIO/NamedRowEvaluator.makeHeaders(arr); 
    ///             X/Cursor.mapResult(eval1(lookups), src1)
    ///         }
    ///         case None => { 
    ///             X/Cursor.finalize!(src1);
    ///             X/Cursor.failure("Cannot find headers")
    ///         }
    ///     }

}
