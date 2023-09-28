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

mod SheetIO.Internal.Apache.Poi.SS.Usermodel {


    pub enum Row(##org.apache.poi.ss.usermodel.Row)


}

mod SheetIO.Internal.Apache.Poi.SS.Usermodel.Row {


    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Row
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Row.{Row}
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Cell
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Cell.{Cell}
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.Sheet.{Sheet}


    pub def getRowNum(row: Row): Int32 = 
        import org.apache.poi.ss.usermodel.Row.getRowNum(): Int32 \ {}; 
        let Row(row1) = row;
        getRowNum(row1)
    
    pub def getCell(ix: Int32, row: Row): Option[Cell] = 
        import org.apache.poi.ss.usermodel.Row.getCell(Int32): ##org.apache.poi.ss.usermodel.Cell \ {}; 
        let Row(row1) = row;
        getCell(row1, ix) |> Object.toOption |> Option.map(Cell)

    pub def getSheet(row: Row): Sheet = 
        import org.apache.poi.ss.usermodel.Row.getSheet(): ##org.apache.poi.ss.usermodel.Sheet \ {};
        let Row(row1) = row;
        getSheet(row1) |> Sheet



    pub def getFirstCellNum(row: Row): Int16 = 
        import org.apache.poi.ss.usermodel.Row.getFirstCellNum(): Int16 \ {};
        let Row(row1) = row;
        getFirstCellNum(row1)


    /// As per Java - index of last cell **plus 1**.
    pub def getLastCellNum(row: Row): Int16 = 
        import org.apache.poi.ss.usermodel.Row.getLastCellNum(): Int16 \ {}; 
        let Row(row1) = row;
        getLastCellNum(row1)


    /// Number of cells used - do not use for trying to find end.
    pub def getPhysicalNumberOfCells(row: Row): Int32 = 
        import org.apache.poi.ss.usermodel.Row.getPhysicalNumberOfCells(): Int32 \ {}; 
        let Row(row1) = row;
        getPhysicalNumberOfCells(row1)
        

    pub def cellIterator(rc: Region[r], row: Row): Iterator[Cell, r, r] \ r = 
        import org.apache.poi.ss.usermodel.Row.cellIterator(): ##java.util.Iterator \ r;
        let Row(row1) = row;
        let iter1 = cellIterator(row1);
        Adaptor.fromIterator(rc, Proxy.Proxy: Proxy[##org.apache.poi.ss.usermodel.Cell], iter1) 
            |> Iterator.map(Cell)

}
