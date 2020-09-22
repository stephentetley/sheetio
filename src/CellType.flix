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

use System.Error.{Generic};

use SheetReader.CellType;
use SheetReader.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString};


namespace SheetReader {

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

}

namespace SheetReader/CellType {

    pub def marshal(x: CellType): ##org.apache.poi.ss.usermodel.CellType = 
        import flix.runtime.spt.sheetreader.Wrapper:get_NONE()      as get_NONE;
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

    pub def unmarshal(x: ##org.apache.poi.ss.usermodel.CellType): CellType = 
        import flix.runtime.spt.sheetreader.Wrapper:get_NONE()      as get_NONE;
        import get org.apache.poi.ss.usermodel.CellType:BLANK       as getBLANK;
        import get org.apache.poi.ss.usermodel.CellType:BOOLEAN     as getBOOLEAN;
        import get org.apache.poi.ss.usermodel.CellType:ERROR       as getERROR;
        import get org.apache.poi.ss.usermodel.CellType:FORMULA     as getFORMULA;
        import get org.apache.poi.ss.usermodel.CellType:NUMERIC     as getNUMERIC;
        import get org.apache.poi.ss.usermodel.CellType:STRING      as getSTRING;
        match x {
            case x1 if x1 == get_NONE()     => TyNone
            case x1 if x1 == getBLANK()     => TyBlank
            case x1 if x1 == getBOOLEAN()   => TyBoolean
            case x1 if x1 == getERROR()     => TyError
            case x1 if x1 == getFORMULA()   => TyFormula
            case x1 if x1 == getNUMERIC()   => TyNumeric
            case x1 if x1 == getSTRING()    => TyString
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
    /// Returns the corresponding day of the week for the integer `n`.
    ///
    /// `n` must be within the range 1 (Monday) to 7 (Sunday) otherwise `Err(_)` is returned.
    ///
    pub def fromInt32(n: Int32): Result[CellType, System.Error] = match n {
        case x1 if x1 == 0      => Ok(TyNone)
        case x1 if x1 == 1      => Ok(TyBlank)
        case x1 if x1 == 2      => Ok(TyBoolean)
        case x1 if x1 == 3      => Ok(TyError)
        case x1 if x1 == 4      => Ok(TyFormula)
        case x1 if x1 == 5      => Ok(TyNumeric)
        case x1 if x1 == 6      => Ok(TyString)
        case _                  => Err(Generic("fromInt32"))
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