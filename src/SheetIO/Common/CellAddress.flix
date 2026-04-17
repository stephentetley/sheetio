/*
 * Copyright 2026 Stephen Tetley
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


pub mod SheetIO.Common.CellAddress {


    import org.apache.poi.ss.util.{CellAddress => JCellAddress}
    import java.lang.{Exception => JException}

    pub enum CellAddress(JCellAddress)

    instance Eq[CellAddress] {
        pub def eq(x: CellAddress, y: CellAddress): Bool = 
            let CellAddress.CellAddress(x1) = x;
            let CellAddress.CellAddress(y1) = y;
            unsafe IO { x1.equals(y1) }
    }

    instance ToString[CellAddress] {
        pub def toString(x: CellAddress): String = 
            let CellAddress.CellAddress(x1) = x;
            unsafe IO { x1.toString() }
    }


    instance Hash[CellAddress] {
        pub def hash(x: CellAddress): Int32 = 
            let CellAddress.CellAddress(x1) = x;
            unsafe IO { x1.hashCode() }
    }


    pub def create(rownum: Int32, colnum: Int32): CellAddress = 
        CellAddress.CellAddress(unsafe IO { new JCellAddress(rownum, colnum) })

    pub def createFromString(addr: String): CellAddress = 
        CellAddress.CellAddress(unsafe IO { new JCellAddress(addr) })

    pub def getColumn(x: CellAddress): Int32 = 
        let CellAddress.CellAddress(x1) = x;
        unsafe IO { x1.getColumn() }

    pub def getRow(x: CellAddress): Int32 = 
        let CellAddress.CellAddress(x1) = x;
        unsafe IO { x1.getRow() }

    pub def formatAsString(x: CellAddress): String = 
        let CellAddress.CellAddress(x1) = x;
        unsafe IO { x1.formatAsString() }

}
