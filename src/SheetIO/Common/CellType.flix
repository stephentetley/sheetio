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


pub mod SheetIO.Common.CellType {
    use SheetIO.Common.CellType


    import org.apache.poi.ss.usermodel.{CellType => JCellType}
    import java.lang.{Exception => JException}

    // Idea - can with live without eagerly translating the Java 
    // enum to a Flix enum and only translate if we need to do 
    // pattern matching?
    // 

    pub enum CellType(JCellType)

    instance Eq[CellType] {
        pub def eq(x: CellType, y: CellType): Bool = 
            let CellType.CellType(x1) = x;
            let CellType.CellType(y1) = y;
            unsafe IO { x1.equals(y1) }
    }

    instance ToString[CellType] {
        pub def toString(x: CellType): String = 
            let CellType.CellType(x1) = x;
            unsafe IO { x1.toString() }
    }


    instance Hash[CellType] {
        pub def hash(x: CellType): Int32 = 
            let CellType.CellType(x1) = x;
            unsafe IO { x1.hashCode() }
    }


    instance ToJava[CellType] {
        type Out = JCellType
        pub def toJava(x: CellType): JCellType = 
            let CellType.CellType(x1) = x;
            x1
    }

    instance ToFlix[CellType] {
        type In = JCellType
        pub def toFlix(x: JCellType): CellType = CellType.CellType(x)
    }


    pub enum CellTypeView {
        case TyNone
        case TyBlank
        case TyBoolean
        case TyError
        case TyFormula
        case TyNumeric
        case TyString
    }
    
    instance ToJava[CellTypeView] {
        type Out = JCellType
        pub def toJava(x: CellTypeView): JCellType = match x {
            case CellTypeView.TyNone => unsafe IO { JCellType._NONE }
            case CellTypeView.TyBlank  => unsafe IO { JCellType.BLANK }
            case CellTypeView.TyBoolean  => unsafe IO { JCellType.BOOLEAN }
            case CellTypeView.TyError  => unsafe IO { JCellType.ERROR }
            case CellTypeView.TyFormula  => unsafe IO { JCellType.FORMULA }
            case CellTypeView.TyNumeric  => unsafe IO { JCellType.NUMERIC }
            case CellTypeView.TyString  => unsafe IO { JCellType.STRING }
        }
    }

    instance ToFlix[CellTypeView] {
        type In = JCellType
        pub def toFlix(x: JCellType): CellTypeView = unsafe IO {
            if (x.equals(JCellType._NONE))          CellTypeView.TyNone
            else if (x.equals(JCellType.BLANK))     CellTypeView.TyBlank
            else if (x.equals(JCellType.BOOLEAN))   CellTypeView.TyBoolean
            else if (x.equals(JCellType.ERROR))     CellTypeView.TyError
            else if (x.equals(JCellType.FORMULA))   CellTypeView.TyFormula
            else if (x.equals(JCellType.NUMERIC))   CellTypeView.TyNumeric
            else if (x.equals(JCellType.STRING))    CellTypeView.TyString
            else bug!("toFlix[CellTypeView]")
        }
    }

    pub def cellTypeView(x: CellType): CellTypeView = 
        let jct = ToJava.toJava(x);
        ToFlix.toFlix(jct)


}
