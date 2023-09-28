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

mod Internal.Apache.Poi.SS.Usermodel {


    pub enum Sheet(##org.apache.poi.ss.usermodel.Sheet)


}

mod Internal.Apache.Poi.SS.Usermodel.Sheet {

    use Internal.Apache.Poi.SS.Usermodel.Sheet
    use Internal.Apache.Poi.SS.Usermodel.Sheet.{Sheet}
    use Internal.Apache.Poi.SS.Usermodel.Row
    use Internal.Apache.Poi.SS.Usermodel.Row.{Row}
    use Internal.Apache.Poi.SS.Usermodel.Workbook
    use Internal.Apache.Poi.SS.Usermodel.Workbook.{Workbook}



    pub def getSheetName(sheet: Sheet): String = 
        import org.apache.poi.ss.usermodel.Sheet.getSheetName(): String \ {}; 
        let Sheet(sheet1) = sheet;
        getSheetName(sheet1)

    pub def getFirstRowNum(sheet: Sheet): Option[Int32] = 
        import org.apache.poi.ss.usermodel.Sheet.getFirstRowNum(): Int32 \ {};
        let Sheet(sheet1) = sheet;
        let ans = getFirstRowNum(sheet1);
        if (ans >= 0)
            ans |> Some
        else
            None

    /// Gets the last row number - this row could be "empty" if cells in the row have been written to but then cleared.
    pub def getLastRowNum(sheet: Sheet): Option[Int32] = 
        import org.apache.poi.ss.usermodel.Sheet.getLastRowNum(): Int32 \ {};
        let Sheet(sheet1) = sheet;
        let ans = getLastRowNum(sheet1);
        if (ans >= 0)
            ans |> Some
        else
            None

    /// Number of rows used - do not use for trying to find end.
    pub def getPhysicalNumberOfRows(sheet: Sheet): Int32 = 
        import org.apache.poi.ss.usermodel.Sheet.getPhysicalNumberOfRows(): Int32 \ {}; 
        let Sheet(sheet1) = sheet;
        getPhysicalNumberOfRows(sheet1)

    pub def getWorkbook(sheet: Sheet): Workbook = 
        import org.apache.poi.ss.usermodel.Sheet.getWorkbook(): ##org.apache.poi.ss.usermodel.Workbook \ {};
        let Sheet(sheet1) = sheet;
        getWorkbook(sheet1) |> Workbook


    pub def getRow(ix: Int32, sheet: Sheet): Result[String, Row] = 
        import org.apache.poi.ss.usermodel.Sheet.getRow(Int32): ##org.apache.poi.ss.usermodel.Row \ {}; 
        let Sheet(sheet1) = sheet;
        let ans = getRow(sheet1, ix);
        if (not Object.isNull(ans))
            ans |> Row |> Ok
        else
            Err("getRow - null")

    pub def rowIterator(rc: Region[r], sheet: Sheet): Iterator[Row, r, r] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.rowIterator(): ##java.util.Iterator \ r;
        let Sheet(sheet1) = sheet;
        let iter1 = rowIterator(sheet1);
        Adaptor.fromIterator(rc, Proxy.Proxy: Proxy[##org.apache.poi.ss.usermodel.Row], iter1) 
            |> Iterator.map(Row)

}
