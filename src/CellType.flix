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



namespace SheetIO {

    ///
    /// CellType
    ///
    pub enum CellType {
        case TyNone,
        case TyBlank,
        case TyBoolean,
        case TyError,
        case TyFormula,
        case TyNumeric,
        case TyString
    }


    instance ToString[CellType] {
        def toString(x: CellType): String = match x {
            case TyNone => "TyNone"
            case TyBlank => "TyBlank"
            case TyBoolean => "TyBoolean"
            case TyError => "TyError"
            case TyFormula => "TyFormula"
            case TyNumeric => "TyNumeric"
            case TyString => "TyString"
        }
    }

    instance FromString[CellType] {
        def fromString(x: String): Option[CellType] = match x {
            case "TyNone" => Some(TyNone)
            case "TyBlank" => Some(TyBlank)
            case "TyBoolean" => Some(TyBoolean)
            case "TyError" => Some(TyError)
            case "TyFormula" => Some(TyFormula)
            case "TyNumeric" => Some(TyNumeric)
            case "TyString" => Some(TyString)
            case _ => None
        }
    }

    instance Eq[CellType] {

        def eq(x: CellType, y: CellType): Bool = match (x, y) {
            case (TyNone, TyNone) => true
            case (TyBlank, TyBlank) => true
            case (TyBoolean, TyBoolean) => true
            case (TyError, TyError) => true
            case (TyFormula, TyFormula) => true
            case (TyNumeric, TyNumeric) => true
            case (TyString, TyString) => true
            case (_, _) => false
        }

        def neq(x: CellType, y: CellType): Bool = not (x == y)
    }

}

namespace SheetIO/CellType {

    use SheetIO.CellType;
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};


    pub def marshal(x: CellType): ##org.apache.poi.ss.usermodel.CellType = 
        import flix.runtime.spt.sheetio.Wrapper:get_NONE()      as get_NONE;
        import get org.apache.poi.ss.usermodel.CellType:BLANK       as getBLANK;
        import get org.apache.poi.ss.usermodel.CellType:BOOLEAN     as getBOOLEAN;
        import get org.apache.poi.ss.usermodel.CellType:ERROR       as getERROR;
        import get org.apache.poi.ss.usermodel.CellType:FORMULA     as getFORMULA;
        import get org.apache.poi.ss.usermodel.CellType:NUMERIC     as getNUMERIC;
        import get org.apache.poi.ss.usermodel.CellType:STRING      as getSTRING;
        match x {
            case TyNone     => get_NONE()
            case TyBlank    => getBLANK()
            case TyBoolean  => getBOOLEAN()
            case TyError    => getERROR()
            case TyFormula  => getFORMULA()
            case TyNumeric  => getNUMERIC()
            case TyString   => getSTRING() 
        } as & Pure


    def primEquals(x: ##org.apache.poi.ss.usermodel.CellType, y: ##org.apache.poi.ss.usermodel.CellType): Bool =
        import java.lang.Object.equals(##java.lang.Object);
        let x1 = x as ##java.lang.Object;
        let y1 = y as ##java.lang.Object;
        equals(x1, y1) as & Pure

    pub def unmarshal(x: ##org.apache.poi.ss.usermodel.CellType): CellType = 
        import flix.runtime.spt.sheetio.Wrapper:get_NONE()      as get_NONE;
        import get org.apache.poi.ss.usermodel.CellType:BLANK       as getBLANK;
        import get org.apache.poi.ss.usermodel.CellType:BOOLEAN     as getBOOLEAN;
        import get org.apache.poi.ss.usermodel.CellType:ERROR       as getERROR;
        import get org.apache.poi.ss.usermodel.CellType:FORMULA     as getFORMULA;
        import get org.apache.poi.ss.usermodel.CellType:NUMERIC     as getNUMERIC;
        import get org.apache.poi.ss.usermodel.CellType:STRING      as getSTRING;
        match x {
            case x1 if primEquals(x1, get_NONE())       => TyNone
            case x1 if primEquals(x1, getBLANK())       => TyBlank
            case x1 if primEquals(x1, getBOOLEAN())     => TyBoolean
            case x1 if primEquals(x1, getERROR())       => TyError
            case x1 if primEquals(x1, getFORMULA())     => TyFormula
            case x1 if primEquals(x1, getNUMERIC())     => TyNumeric
            case x1 if primEquals(x1, getSTRING())      => TyString
        } as & Pure


    ///
    /// Returns an integer representation of the CellType.
    ///
    /// The representation follows the ISO-8601 standard.
    /// Days are counted from 1 (Monday) to 7 (Sunday).
    ///
    pub def toInt32(x: CellType): Int32 = match x {
        case TyNone     => 0
        case TyBlank    => 1    
        case TyBoolean  => 2
        case TyError    => 3
        case TyFormula  => 4
        case TyNumeric  => 5
        case TyString   => 6
    }

    ///
    ///
    pub def fromInt32(n: Int32): Option[CellType] = match n {
        case x1 if x1 == 0      => Some(TyNone)
        case x1 if x1 == 1      => Some(TyBlank)
        case x1 if x1 == 2      => Some(TyBoolean)
        case x1 if x1 == 3      => Some(TyError)
        case x1 if x1 == 4      => Some(TyFormula)
        case x1 if x1 == 5      => Some(TyNumeric)
        case x1 if x1 == 6      => Some(TyString)
        case _                  => None
    }

    pub def equals(x: CellType, y: CellType): Bool = toInt32(x) == toInt32(y)

    pub def compare(x: CellType, y: CellType): Int32 = Int32.compare(toInt32(x), toInt32(y))


    pub def format (x: CellType): String = match x {
        case TyNone     => "_NONE"
        case TyBlank    => "BLANK"
        case TyBoolean  => "BOOLEAN"
        case TyError    => "ERROR"
        case TyFormula  => "FORMULA"
        case TyNumeric  => "NUMERIC"
        case TyString   => "STRING"
    }


}