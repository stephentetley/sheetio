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


pub mod SheetIO.Usermodel {

    import org.apache.poi.ss.usermodel.{Workbook => JWorkbook}
    import org.apache.poi.ss.usermodel.{Sheet => JSheet}
    import org.apache.poi.ss.usermodel.{Row => JRow}
    import org.apache.poi.ss.usermodel.{Cell => JCell}
    import org.apache.poi.ss.usermodel.{Table => JTable}
    import org.apache.poi.ss.usermodel.{Row$MissingCellPolicy => JRowMissingCellPolicy}
    import org.apache.poi.ss.usermodel.{CellStyle => JCellStyle}
    import org.apache.poi.ss.usermodel.{FillPatternType => JFillPatternType}

    pub enum Workbook(JWorkbook)

    instance ToJava[Workbook] {
        type Out = JWorkbook
        pub def toJava(x: Workbook): JWorkbook =
            let Workbook.Workbook(x1) = x;
            x1
    }

    instance ToFlix[Workbook] {
        type In = JWorkbook
        pub def toFlix(x: JWorkbook): Workbook = Workbook.Workbook(x)
    }

    pub enum Sheet(JSheet)

    instance ToJava[Sheet] {
        type Out = JSheet
        pub def toJava(x: Sheet): JSheet =
            let Sheet.Sheet(x1) = x;
            x1
    }

    instance ToFlix[Sheet] {
        type In = JSheet
        pub def toFlix(x: JSheet): Sheet = Sheet.Sheet(x)
    }

    pub enum Row(JRow)
    
    instance ToJava[Row] {
        type Out = JRow
        pub def toJava(x: Row): JRow = 
            let Row.Row(x1) = x;
            x1
    }

    instance ToFlix[Row] {
        type In = JRow
        pub def toFlix(x: JRow): Row = Row.Row(x)
    }


    pub enum Cell(JCell)

    instance ToJava[Cell] {
        type Out = JCell
        pub def toJava(x: Cell): JCell = 
            let Cell.Cell(x1) = x;
            x1
    }

    instance ToFlix[Cell] {
        type In = JCell
        pub def toFlix(x: JCell): Cell = Cell.Cell(x)
    }


    
    pub enum Table(JTable)

    instance ToJava[Table] {
        type Out = JTable
        pub def toJava(x: Table): JTable = 
            let Table.Table(x1) = x;
            x1
    }

    instance ToFlix[Table] {
        type In = JTable
        pub def toFlix(x: JTable): Table = Table.Table(x)
    }    



    // TODO - maybe we don't want to support this as we don't want 
    // to be introducing nulls
    pub enum RowMissingCellPolicy with Eq, ToString {
        case CREATE_NULL_AS_BLANK 
        case RETURN_BLANK_AS_NULL 
        case RETURN_NULL_AND_BLANK 
    }


    
    instance ToJava[RowMissingCellPolicy] {
        type Out = JRowMissingCellPolicy
        pub def toJava(x: RowMissingCellPolicy): JRowMissingCellPolicy = match x {
            case RowMissingCellPolicy.CREATE_NULL_AS_BLANK      => unsafe IO { JRowMissingCellPolicy.CREATE_NULL_AS_BLANK }
            case RowMissingCellPolicy.RETURN_BLANK_AS_NULL      => unsafe IO { JRowMissingCellPolicy.RETURN_BLANK_AS_NULL }
            case RowMissingCellPolicy.RETURN_NULL_AND_BLANK     => unsafe IO { JRowMissingCellPolicy.RETURN_NULL_AND_BLANK }
        }
    }

    instance ToFlix[RowMissingCellPolicy] {
        type In = JRowMissingCellPolicy
        pub def toFlix(x: JRowMissingCellPolicy): RowMissingCellPolicy = unsafe IO {
            if (x.equals(JRowMissingCellPolicy.CREATE_NULL_AS_BLANK))           RowMissingCellPolicy.CREATE_NULL_AS_BLANK
            else if (x.equals(JRowMissingCellPolicy.RETURN_BLANK_AS_NULL))      RowMissingCellPolicy.RETURN_BLANK_AS_NULL
            else if (x.equals(JRowMissingCellPolicy.RETURN_NULL_AND_BLANK))     RowMissingCellPolicy.RETURN_NULL_AND_BLANK
            else bug!("toFlix[RowMissingCellPolicy]")
        }
    }

    pub enum CellStyle(JCellStyle)

    instance ToJava[CellStyle] {
        type Out = JCellStyle
        pub def toJava(x: CellStyle): JCellStyle =
            let CellStyle.CellStyle(x1) = x;
            x1
    }

    instance ToFlix[CellStyle] {
        type In = JCellStyle
        pub def toFlix(x: JCellStyle): CellStyle = CellStyle.CellStyle(x)
    }

    pub enum FillPatternType with Eq, ToString {
        case ALT_BARS
        case BIG_SPOTS
        case BRICKS
        case DIAMONDS
        case FINE_DOTS
        case LEAST_DOTS
        case LESS_DOTS
        case NO_FILL
        case SOLID_FOREGROUND
        case SPARSE_DOTS
        case SQUARES
        case THICK_BACKWARD_DIAG
        case THICK_FORWARD_DIAG
        case THICK_HORZ_BANDS
        case THICK_VERT_BANDS
        case THIN_BACKWARD_DIAG
        case THIN_FORWARD_DIAG
        case THIN_HORZ_BANDS
        case THIN_VERT_BANDS
    }

    instance ToJava[FillPatternType] {
        type Out = JFillPatternType
        pub def toJava(x: FillPatternType): JFillPatternType = match x {
            case FillPatternType.ALT_BARS               => unsafe IO { JFillPatternType.ALT_BARS }
            case FillPatternType.BIG_SPOTS              => unsafe IO { JFillPatternType.BIG_SPOTS }
            case FillPatternType.BRICKS                 => unsafe IO { JFillPatternType.BRICKS }
            case FillPatternType.DIAMONDS               => unsafe IO { JFillPatternType.DIAMONDS }
            case FillPatternType.FINE_DOTS              => unsafe IO { JFillPatternType.FINE_DOTS }
            case FillPatternType.LEAST_DOTS             => unsafe IO { JFillPatternType.LEAST_DOTS }
            case FillPatternType.LESS_DOTS              => unsafe IO { JFillPatternType.LESS_DOTS }
            case FillPatternType.NO_FILL                => unsafe IO { JFillPatternType.NO_FILL }
            case FillPatternType.SOLID_FOREGROUND       => unsafe IO { JFillPatternType.SOLID_FOREGROUND }
            case FillPatternType.SPARSE_DOTS            => unsafe IO { JFillPatternType.SPARSE_DOTS }
            case FillPatternType.SQUARES                => unsafe IO { JFillPatternType.SQUARES }
            case FillPatternType.THICK_BACKWARD_DIAG    => unsafe IO { JFillPatternType.THICK_BACKWARD_DIAG }
            case FillPatternType.THICK_FORWARD_DIAG     => unsafe IO { JFillPatternType.THICK_FORWARD_DIAG }
            case FillPatternType.THICK_HORZ_BANDS       => unsafe IO { JFillPatternType.THICK_HORZ_BANDS }
            case FillPatternType.THICK_VERT_BANDS       => unsafe IO { JFillPatternType.THICK_VERT_BANDS }
            case FillPatternType.THIN_BACKWARD_DIAG     => unsafe IO { JFillPatternType.THIN_BACKWARD_DIAG }
            case FillPatternType.THIN_FORWARD_DIAG      => unsafe IO { JFillPatternType.THIN_FORWARD_DIAG }
            case FillPatternType.THIN_HORZ_BANDS        => unsafe IO { JFillPatternType.THIN_HORZ_BANDS }
            case FillPatternType.THIN_VERT_BANDS        => unsafe IO { JFillPatternType.THIN_VERT_BANDS }
        }
    }


    instance ToFlix[FillPatternType] {
        type In = JFillPatternType
        pub def toFlix(x: JFillPatternType): FillPatternType = unsafe IO {
            if (x.equals(JFillPatternType.ALT_BARS))                    FillPatternType.ALT_BARS
            else if (x.equals(JFillPatternType.BIG_SPOTS))              FillPatternType.BIG_SPOTS
            else if (x.equals(JFillPatternType.BRICKS))                 FillPatternType.BRICKS
            else if (x.equals(JFillPatternType.DIAMONDS))               FillPatternType.DIAMONDS
            else if (x.equals(JFillPatternType.FINE_DOTS))              FillPatternType.FINE_DOTS
            else if (x.equals(JFillPatternType.LEAST_DOTS))             FillPatternType.LEAST_DOTS
            else if (x.equals(JFillPatternType.LESS_DOTS))              FillPatternType.LESS_DOTS
            else if (x.equals(JFillPatternType.NO_FILL))                FillPatternType.NO_FILL
            else if (x.equals(JFillPatternType.SOLID_FOREGROUND))       FillPatternType.SOLID_FOREGROUND
            else if (x.equals(JFillPatternType.SPARSE_DOTS))            FillPatternType.SPARSE_DOTS
            else if (x.equals(JFillPatternType.SQUARES))                FillPatternType.SQUARES
            else if (x.equals(JFillPatternType.THICK_BACKWARD_DIAG))    FillPatternType.THICK_BACKWARD_DIAG
            else if (x.equals(JFillPatternType.THICK_FORWARD_DIAG))     FillPatternType.THICK_FORWARD_DIAG
            else if (x.equals(JFillPatternType.THICK_HORZ_BANDS))       FillPatternType.THICK_HORZ_BANDS
            else if (x.equals(JFillPatternType.THICK_VERT_BANDS))       FillPatternType.THICK_VERT_BANDS
            else if (x.equals(JFillPatternType.THIN_BACKWARD_DIAG))     FillPatternType.THIN_BACKWARD_DIAG
            else if (x.equals(JFillPatternType.THIN_FORWARD_DIAG))      FillPatternType.THIN_FORWARD_DIAG
            else if (x.equals(JFillPatternType.THIN_HORZ_BANDS))        FillPatternType.THIN_HORZ_BANDS
            else if (x.equals(JFillPatternType.THIN_VERT_BANDS))        FillPatternType.THIN_VERT_BANDS
            else bug!("toFlix[FillPatternType]")
        }
    }

}

