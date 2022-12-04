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

namespace Interop/Apache/Poi/SS/Usermodel {

    /// Mutable / Interface

    pub enum Row[_: Region](##org.apache.poi.ss.usermodel.Row)

    instance Scoped[Row] {
        pub def regionOf(_: Row[r]): Region[r] = unsafe_cast () as Region[r]
    }
}

namespace Interop/Apache/Poi/SS/Usermodel/Row {


    use Interop/Apache/Poi/SS/Usermodel.Row;
    use Interop/Apache/Poi/SS/Usermodel.Row.{Row};
    use Interop/Apache/Poi/SS/Usermodel.Cell;
    use Interop/Apache/Poi/SS/Usermodel.Cell.{Cell};
    use Interop/Apache/Poi/SS/Usermodel.Sheet;
    use Interop/Apache/Poi/SS/Usermodel.Sheet.{Sheet};

    pub def getHH(row: Row[r]): ##org.apache.poi.ss.usermodel.Row =
        let Row(row1) = row;
        row1

    
    pub def wrapHH(_: Region[r], prim:  ##org.apache.poi.ss.usermodel.Row): Row[r] =
        Row(prim)

    pub def getRowNum(row: Row[r]): Int32 \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getRowNum(): Int32 \ r; 
        let Row(row1) = row;
        getRowNum(row1)
    
    pub def getCell(ix: Int32, row: Row[r]): Option[Cell[r]] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getCell(Int32): ##org.apache.poi.ss.usermodel.Cell \ r; 
        let Row(row1) = row;
        getCell(row1, ix) |> Object.toOption |> Option.map(Cell)

    pub def getSheet(row: Row[r]): Sheet[r] \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getSheet(): ##org.apache.poi.ss.usermodel.Sheet \ r;
        let Row(row1) = row;
        getSheet(row1) |> Sheet

    pub def createCell!(ix: Int32, row: Row[r]): Result[Cell[r], String] \ Write(r) = 
        try {
            import org.apache.poi.ss.usermodel.Row.createCell(Int32): ##org.apache.poi.ss.usermodel.Cell \ r;
            let Row(row1) = row;
            createCell(row1, ix) |> Cell |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


    pub def getFirstCellNum(row: Row[r]): Int16 \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getFirstCellNum(): Int16 \ r;
        let Row(row1) = row;
        getFirstCellNum(row1)


    /// As per Java - index of last cell **plus 1**.
    pub def getLastCellNum(row: Row[r]): Int16 \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getLastCellNum(): Int16 \ r; 
        let Row(row1) = row;
        getLastCellNum(row1)


    /// Number of cells used - do not use for trying to find end.
    pub def getPhysicalNumberOfCells(row: Row[r]): Int32 \ Read(r) = 
        import org.apache.poi.ss.usermodel.Row.getPhysicalNumberOfCells(): Int32 \ r; 
        let Row(row1) = row;
        getPhysicalNumberOfCells(row1)
        

}
