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


pub mod SheetIO.Common.CellReference {


    import org.apache.poi.ss.util.{CellReference => JCellReference}
    import java.lang.{Exception => JException}

    pub enum CellReference(JCellReference)

    instance Eq[CellReference] {
        pub def eq(x: CellReference, y: CellReference): Bool = 
            let CellReference.CellReference(x1) = x;
            let CellReference.CellReference(y1) = y;
            unsafe IO { x1.equals(y1) }
    }

    instance ToString[CellReference] {
        pub def toString(x: CellReference): String = 
            let CellReference.CellReference(x1) = x;
            unsafe IO { x1.toString() }
    }


    instance Hash[CellReference] {
        pub def hash(x: CellReference): Int32 = 
            let CellReference.CellReference(x1) = x;
            unsafe IO { x1.hashCode() }
    }


    instance ToJava[CellReference] {
        type Out = JCellReference
        pub def toJava(x: CellReference): JCellReference = 
            let CellReference.CellReference(x1) = x;
            x1
    }

    instance ToFlix[CellReference] {
        type In = JCellReference
        pub def toFlix(x: JCellReference): CellReference = CellReference.CellReference(x)
    }

    // Probably should be Result[err, CellReference]...
    pub def createFromString(addr: String): CellReference = 
        CellReference.CellReference(unsafe IO { new JCellReference(addr) })

    pub def formatAsString(cr: CellReference): String = 
        let cr1 = ToJava.toJava(cr);
        unsafe IO { cr1.formatAsString() }

    pub def getSheetName(cr: CellReference): Option[String] = 
        let cr1 = ToJava.toJava(cr);
        let ans = unsafe IO { cr1.getSheetName() };
        if (not Object.isNull(ans)) Some(ans) else None

    pub def getCol(cr: CellReference): Int16 = 
        let cr1 = ToJava.toJava(cr);
        unsafe IO { cr1.getCol() }

    pub def getRow(cr: CellReference): Int32 = 
        let cr1 = ToJava.toJava(cr);
        unsafe IO { cr1.getRow() }

    pub def isRowAbsolute(cr: CellReference): Bool = 
        let cr1 = ToJava.toJava(cr);
        unsafe IO { cr1.isRowAbsolute() }

    pub def isColAbsolute(cr: CellReference): Bool = 
        let cr1 = ToJava.toJava(cr);
        unsafe IO { cr1.isColAbsolute() }

    // NOTE - ZERO BASED!

    pub def convertNumToColString(i: Int32): String = 
        unsafe IO { JCellReference.convertNumToColString(i) }

}
