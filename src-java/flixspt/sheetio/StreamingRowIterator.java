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

package flixspt.sheetio;

import java.io.IOException;
import java.nio.file.Path;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;

/// This is a properly streaming (i.e. low memory residency) cursor using the
/// pjfanning's fork of excel-streaming-reader [original: com.monitorjbl.xlsx.StreamingReader]
/// The pjfanning fork works with poi-5.0.0

public class StreamingRowIterator {

    private final Workbook workbook;
    private final Iterator<Row> iter;

    protected StreamingRowIterator(Workbook wb, String sheetName) {
        workbook = wb;
        iter = workbook.getSheet(sheetName).rowIterator();
    }

    public static StreamingRowIterator createIteratorWithSheetName(Path path, String sheetName) throws Exception {
        InputStream instr = new FileInputStream(path.toFile());
        Workbook wb = StreamingReader.builder().open(instr);
        return new StreamingRowIterator(wb, sheetName);
    }

    public static StreamingRowIterator createIteratorWithSheetNumber(Path path, int sheetNumber) throws Exception {
        InputStream instr = new FileInputStream(path.toFile());
        Workbook wb = StreamingReader.builder().open(instr);
        String sheetName = wb.getSheetName(sheetNumber);
        return new StreamingRowIterator(wb, sheetName);
    }

    public boolean hasNext() {
        return iter.hasNext();
    }

    public Row next() throws Exception {
        return iter.next();
    }

    public void close() throws IOException {
        workbook.close();
    }

}
