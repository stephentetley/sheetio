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


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;


/* Note - this class iterates the rows of org.apache.poi.ss.usermodel.Sheet.
 * This is only because delegating to Java to iterating rows was found to
 * be a reliable method of reading up to the end of the sheet. Trying to
 * find the end with a call to `getLastRowNum` in Flix code seemed error
 * prone and was abandoned as a strategy.
 *
 * As far as I am aware the whole Sheet is read into memory and it
 * is streamed while in memory - this module does not implement
 * streaming as a way of achieving low memory usage.
*/
public class POIRowIterator {

    private Iterator<Row> iter;

    public POIRowIterator(Sheet sheet) throws Exception {
        this.iter = sheet.rowIterator();
    }


    public boolean hasNext() {
        return this.iter.hasNext();
    }

    public Row next() throws Exception { return this.iter.next(); }


}
