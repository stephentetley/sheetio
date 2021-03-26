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

namespace SheetIO/StreamingReader {

    use Data.Stream;
    use System.FilePath.{FilePath};

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};


    type alias POIRow = ##org.apache.poi.ss.usermodel.Row

    ///
    /// Returns a new Row Stream.
    ///
    pub def newRowStream(src: System.FilePath, sheetName: String): Result[Stream[POIRow], String] & Impure =
        import new flix.runtime.spt.sheetio.StreamingRowIterator(##java.nio.file.Path, String) as newRowIterator;
        import flix.runtime.spt.sheetio.StreamingRowIterator.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.next() as ffiNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.close() as ffiClose;
        try {
            let FilePath(src1) = src;
            let iter = newRowIterator(src1, sheetName);
            let streamiter =  {
                  hasNext = () -> ffiHasNext(iter)
                , next = () -> ffiNext(iter)
                , finalize = () -> ffiClose(iter)
            };
            Ok(Data/Stream.fromIterator(streamiter))
        } catch {
            case _: ##java.lang.Exception => Err("newRowStream")
        }
    ///
    /// Returns a new Row Stream for sheet at `ix`.
    ///
    pub def newRowStreamAt(src: System.FilePath, sheetIndex: Int32): Result[Stream[POIRow], String] & Impure =
        import new flix.runtime.spt.sheetio.StreamingRowIterator(##java.nio.file.Path, Int32) as newRowIterator;
        import flix.runtime.spt.sheetio.StreamingRowIterator.hasNext() as ffiHasNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.next() as ffiNext;
        import flix.runtime.spt.sheetio.StreamingRowIterator.close() as ffiClose;
        try {
            let FilePath(src1) = src;
            let iter = newRowIterator(src1, sheetIndex);
            let streamiter =  {
                  hasNext = () -> ffiHasNext(iter)
                , next = () -> ffiNext(iter)
                , finalize = () -> ffiClose(iter)
            };
            Ok(Data/Stream.fromIterator(streamiter))
        } catch {
            case _: ##java.lang.Exception => Err("newRowStreamAt")
        }



}