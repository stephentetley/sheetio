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

    use EffectExtras.Fail.Fail
    use JavaInterop.AdaptorChecked
    use JavaInterop.AdaptorUnchecked

    use SheetIO.Common.CellAddress

    import org.apache.poi.ss.util.{CellReference => JCellReference}
    import org.apache.poi.ss.util.{CellAddress => JCellAddress}
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


    ///
    /// A trait for types that can be converted to CellReferences.
    ///
    pub trait ToCellReference[a] {
        ///
        /// Returns a CellReference from the given x.
        ///
        pub def toCellReference(x: a): Option[CellReference]
    }

    instance ToCellReference[String] {
        pub def toCellReference(x: String): Option[CellReference] = 
            try {
                let ans = unsafe IO { new JCellReference(x) };
                Some(CellReference.CellReference(ans))
            } catch {
                case _: JException => None
            }
    }

    instance ToCellReference[CellReference] {
        pub def toCellReference(x: CellReference): Option[CellReference] = 
            Some(x)
    }


    // Probably should be Result[err, CellReference]...
    pub def createFromString(addr: String): CellReference = 
        CellReference.CellReference(unsafe IO { new JCellReference(addr) })

    pub def formatAsString(x: CellReference): String = 
        AdaptorUnchecked.extract(x, x1 -> unsafe IO { x1.formatAsString() })

    pub def getSheetName(x: CellReference): String \ Fail = 
        AdaptorChecked.extract(x, x1 -> unsafe IO { x1.getSheetName() })

    pub def getCol(x: CellReference): Int16 = 
        AdaptorUnchecked.extract(x, x1 -> unsafe IO { x1.getCol() })

    pub def getRow(x: CellReference): Int32 = 
        AdaptorUnchecked.extract(x, x1 -> unsafe IO { x1.getRow() })

    pub def isRowAbsolute(x: CellReference): Bool = 
        AdaptorUnchecked.extract(x, x1 -> unsafe IO { x1.isRowAbsolute() })

    pub def isColAbsolute(x: CellReference): Bool = 
        AdaptorUnchecked.extract(x, x1 -> unsafe IO { x1.isColAbsolute() })

    // NOTE - ZERO BASED!

    pub def convertNumToColString(i: Int32): String = 
        unsafe IO { JCellReference.convertNumToColString(i) }

    pub def toCellAddress(x: CellReference): CellAddress = 
        AdaptorUnchecked.op(x, x1 -> unsafe IO { new JCellAddress(x1) })


}
