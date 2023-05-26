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

mod Interop.Apache.Poi.SS.Usermodel {

    /// Mutable / Interface

    pub enum Sheet[_: Region](##org.apache.poi.ss.usermodel.Sheet)


}

mod Interop.Apache.Poi.SS.Usermodel.Sheet {

    use Interop.Apache.Poi.SS.Usermodel.Sheet
    use Interop.Apache.Poi.SS.Usermodel.Sheet.{Sheet}
    use Interop.Apache.Poi.SS.Usermodel.Row
    use Interop.Apache.Poi.SS.Usermodel.Row.{Row}
    use Interop.Apache.Poi.SS.Usermodel.Workbook
    use Interop.Apache.Poi.SS.Usermodel.Workbook.{Workbook}

    pub def getHH(sheet: Sheet[r]): ##org.apache.poi.ss.usermodel.Sheet =
        let Sheet(sheet1) = sheet;
        sheet1
   
    pub def wrapHH(_: Region[r], s: ##org.apache.poi.ss.usermodel.Sheet): Sheet[r] =
        Sheet(s)


    pub def getSheetName(sheet: Sheet[r]): String \ r = 
        import org.apache.poi.ss.usermodel.Sheet.getSheetName(): String \ r; 
        let Sheet(sheet1) = sheet;
        getSheetName(sheet1)

    pub def getFirstRowNum(sheet: Sheet[r]): Option[Int32] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.getFirstRowNum(): Int32 \ r;
        let Sheet(sheet1) = sheet;
        let ans = getFirstRowNum(sheet1);
        if (ans >= 0)
            ans |> Some
        else
            None

    /// Gets the last row number - this row could be "empty" if cells in the row have been written to but then cleared.
    pub def getLastRowNum(sheet: Sheet[r]): Option[Int32] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.getLastRowNum(): Int32 \ r;
        let Sheet(sheet1) = sheet;
        let ans = getLastRowNum(sheet1);
        if (ans >= 0)
            ans |> Some
        else
            None

    /// Number of rows used - do not use for trying to find end.
    pub def getPhysicalNumberOfRows(sheet: Sheet[r]): Int32 \ r = 
        import org.apache.poi.ss.usermodel.Sheet.getPhysicalNumberOfRows(): Int32 \ r; 
        let Sheet(sheet1) = sheet;
        getPhysicalNumberOfRows(sheet1)

    pub def getWorkbook(sheet: Sheet[r]): Workbook[r] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.getWorkbook(): ##org.apache.poi.ss.usermodel.Workbook \ r;
        let Sheet(sheet1) = sheet;
        getWorkbook(sheet1) |> Workbook


    pub def getRow(ix: Int32, sheet: Sheet[r]): Result[String, Row[r]] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.getRow(Int32): ##org.apache.poi.ss.usermodel.Row \ r; 
        let Sheet(sheet1) = sheet;
        let ans = getRow(sheet1, ix);
        if (not Object.isNull(ans))
            ans |> Row |> Ok
        else
            Err("getRow - null")

    pub def createRow!(ix: Int32, sheet: Sheet[r]): Row[r] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.createRow(Int32): ##org.apache.poi.ss.usermodel.Row \ r;
        let Sheet(sheet1) = sheet; 
        createRow(sheet1, ix) |> Row

    pub def rowIterator(rc: Region[r1], sheet: Sheet[r]): Iterator[Row[r1], r + r1, r1] \ r = 
        import org.apache.poi.ss.usermodel.Sheet.rowIterator(): ##java.util.Iterator \ r;
        let ocast = obj -> Interop.Apache.Poi.SS.Usermodel.Row.wrapHH(rc, unchecked_cast(obj as ##org.apache.poi.ss.usermodel.Row));
        let Sheet(sheet1) = sheet;
        let iterHH = rowIterator(sheet1);
        Interop.Util.IteratorZ.wrapHH(rc, iterHH) 
            |> Interop.Util.IteratorZ.toFlixIterator(rc, ocast)

}
