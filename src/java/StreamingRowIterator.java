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

package flix.runtime.spt.sheetio;

import java.io.IOException;
import java.nio.file.Path;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;

public class StreamingRowIterator {
    private Workbook wb = null;
    private Iterator<Row> rowIterator;

    public StreamingRowIterator(Path path, String sheetName) throws Exception {
        InputStream instr = new FileInputStream(path.toFile());
        wb = StreamingReader.builder().open(instr);

        rowIterator = wb.getSheet(sheetName).rowIterator();

    }

    public StreamingRowIterator(Path path, int sheetnumber) throws Exception {
        InputStream instr = new FileInputStream(path.toFile());
        wb = StreamingReader.builder().open(instr);

        rowIterator = wb.getSheetAt(sheetnumber).rowIterator();

    }

    public boolean hasNext() {
        return rowIterator.hasNext();
    }

    public Row next() throws Exception {
        Row row = rowIterator.next();
        return row;
    }

    public void close() throws IOException {
        wb.close();
    }

}
