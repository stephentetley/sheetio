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



mod SheetIO {

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}
    
    ///
    /// CellType
    ///
    pub enum CellType with Eq, Order, ToString{
        case TyNone,
        case TyBlank,
        case TyBoolean,
        case TyError,
        case TyFormula,
        case TyNumeric,
        case TyString
    }


    instance FromString[CellType] {
        pub def fromString(x: String): Option[CellType] = match x {
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

    instance Hash[CellType] {
        pub def hash(x: CellType): Int32 = SheetIO.CellType.hashCode(x)
    }

}

mod SheetIO.CellType {

    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}


    pub def toNative(x: CellType): ##org.apache.poi.ss.usermodel.CellType = 
        import static get org.apache.poi.ss.usermodel.CellType._NONE: ##org.apache.poi.ss.usermodel.CellType \ {} as get_NONE;
        import static get org.apache.poi.ss.usermodel.CellType.BLANK: ##org.apache.poi.ss.usermodel.CellType \ {} as getBLANK;
        import static get org.apache.poi.ss.usermodel.CellType.BOOLEAN: ##org.apache.poi.ss.usermodel.CellType \ {} as getBOOLEAN;
        import static get org.apache.poi.ss.usermodel.CellType.ERROR: ##org.apache.poi.ss.usermodel.CellType \ {} as getERROR;
        import static get org.apache.poi.ss.usermodel.CellType.FORMULA: ##org.apache.poi.ss.usermodel.CellType \ {} as getFORMULA;
        import static get org.apache.poi.ss.usermodel.CellType.NUMERIC: ##org.apache.poi.ss.usermodel.CellType \ {} as getNUMERIC;
        import static get org.apache.poi.ss.usermodel.CellType.STRING: ##org.apache.poi.ss.usermodel.CellType \ {} as getSTRING;
        match x {
            case TyNone     => get_NONE()
            case TyBlank    => getBLANK()
            case TyBoolean  => getBOOLEAN()
            case TyError    => getERROR()
            case TyFormula  => getFORMULA()
            case TyNumeric  => getNUMERIC()
            case TyString   => getSTRING() 
        }


    def primEquals(x: ##org.apache.poi.ss.usermodel.CellType, y: ##org.apache.poi.ss.usermodel.CellType): Bool =
        import java.lang.Object.equals(##java.lang.Object): Bool \ {};
        let x1 = checked_cast(x);
        let y1 = checked_cast(y);
        equals(x1, y1)

    pub def fromNative(x: SheetIO.Internal.Apache.Poi.SS.Usermodel.CellType): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType._NONE: ##org.apache.poi.ss.usermodel.CellType \ {} as get_NONE;
        import static get org.apache.poi.ss.usermodel.CellType.BLANK: ##org.apache.poi.ss.usermodel.CellType \ {} as getBLANK;
        import static get org.apache.poi.ss.usermodel.CellType.BOOLEAN: ##org.apache.poi.ss.usermodel.CellType \ {} as getBOOLEAN;
        import static get org.apache.poi.ss.usermodel.CellType.ERROR: ##org.apache.poi.ss.usermodel.CellType \ {} as getERROR;
        import static get org.apache.poi.ss.usermodel.CellType.FORMULA: ##org.apache.poi.ss.usermodel.CellType \ {} as getFORMULA;
        import static get org.apache.poi.ss.usermodel.CellType.NUMERIC: ##org.apache.poi.ss.usermodel.CellType \ {} as getNUMERIC;
        import static get org.apache.poi.ss.usermodel.CellType.STRING: ##org.apache.poi.ss.usermodel.CellType \ {} as getSTRING;
        let SheetIO.Internal.Apache.Poi.SS.Usermodel.CellType.CellType(x1) = x;
        match x1 {
            case x2 if primEquals(x2, get_NONE())       => TyNone
            case x2 if primEquals(x2, getBLANK())       => TyBlank
            case x2 if primEquals(x2, getBOOLEAN())     => TyBoolean
            case x2 if primEquals(x2, getERROR())       => TyError
            case x2 if primEquals(x2, getFORMULA())     => TyFormula
            case x2 if primEquals(x2, getNUMERIC())     => TyNumeric
            case x2 if primEquals(x2, getSTRING())      => TyString
            case _                                      => bug!("fromNative")
        }


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

    pub def hashCode(x: CellType): Int32 =
        import java.lang.Enum.hashCode(): Int32 \ {};
        let x1 = unchecked_cast(toNative(x) as ##java.lang.Enum);
        hashCode(x1)


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