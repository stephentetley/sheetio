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
    import org.apache.poi.ss.usermodel.{Comment => JComment}
    import org.apache.poi.ss.usermodel.{Table => JTable}

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

    pub enum Comment(JComment)

    instance ToJava[Comment] {
        type Out = JComment
        pub def toJava(x: Comment): JComment = 
            let Comment.Comment(x1) = x;
            x1
    }

    instance ToFlix[Comment] {
        type In = JComment
        pub def toFlix(x: JComment): Comment = Comment.Comment(x)
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


}
