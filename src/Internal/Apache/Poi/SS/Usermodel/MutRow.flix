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

    /// Mutable / Interface

    pub enum MutRow[_: Region](##org.apache.poi.ss.usermodel.Row)


}

mod Internal.Apache.Poi.SS.Usermodel.MutRow {


    use Internal.Apache.Poi.SS.Usermodel.MutRow
    use Internal.Apache.Poi.SS.Usermodel.MutRow.{MutRow}
    use Internal.Apache.Poi.SS.Usermodel.MutCell
    use Internal.Apache.Poi.SS.Usermodel.MutCell.{MutCell}
    use Internal.Apache.Poi.SS.Usermodel.Sheet
    use Internal.Apache.Poi.SS.Usermodel.Sheet.{Sheet}

    pub def getHH(row: MutRow[r]): ##org.apache.poi.ss.usermodel.Row =
        let MutRow(row1) = row;
        row1

    
    pub def wrapHH(_: Region[r], prim:  ##org.apache.poi.ss.usermodel.Row): MutRow[r] =
        MutRow(prim)

    pub def getRowNum(row: MutRow[r]): Int32 \ r = 
        import org.apache.poi.ss.usermodel.Row.getRowNum(): Int32 \ r; 
        let MutRow(row1) = row;
        getRowNum(row1)
    
    pub def getCell(ix: Int32, row: MutRow[r]): Option[MutCell[r]] \ r = 
        import org.apache.poi.ss.usermodel.Row.getCell(Int32): ##org.apache.poi.ss.usermodel.Cell \ r; 
        let MutRow(row1) = row;
        getCell(row1, ix) |> Object.toOption |> Option.map(MutCell)

    pub def getSheet(row: MutRow[r]): Sheet \ r = 
        import org.apache.poi.ss.usermodel.Row.getSheet(): ##org.apache.poi.ss.usermodel.Sheet \ r;
        let MutRow(row1) = row;
        getSheet(row1) |> Sheet

    pub def createCell!(ix: Int32, row: MutRow[r]): Result[String, MutCell[r]] \ r = 
        Result.tryCatch(_ -> {
            import org.apache.poi.ss.usermodel.Row.createCell(Int32): ##org.apache.poi.ss.usermodel.Cell \ r;
            let MutRow(row1) = row;
            createCell(row1, ix) |> MutCell
        })


    pub def getFirstCellNum(row: MutRow[r]): Int16 \ r = 
        import org.apache.poi.ss.usermodel.Row.getFirstCellNum(): Int16 \ r;
        let MutRow(row1) = row;
        getFirstCellNum(row1)


    /// As per Java - index of last cell **plus 1**.
    pub def getLastCellNum(row: MutRow[r]): Int16 \ r = 
        import org.apache.poi.ss.usermodel.Row.getLastCellNum(): Int16 \ r; 
        let MutRow(row1) = row;
        getLastCellNum(row1)


    /// Number of cells used - do not use for trying to find end.
    pub def getPhysicalNumberOfCells(row: MutRow[r]): Int32 \ r = 
        import org.apache.poi.ss.usermodel.Row.getPhysicalNumberOfCells(): Int32 \ r; 
        let MutRow(row1) = row;
        getPhysicalNumberOfCells(row1)
        

}
