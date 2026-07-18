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


pub mod SheetIO.Usermodel.Table {
    
    
    
    use EffectExtras.Alt.Alt
    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.{Cell}

    import org.apache.poi.ss.usermodel.{Table => JTable}
    import java.lang.{Exception => JException}


    pub eff Table {
        def contains(cell: Cell): Bool
        def getName(): String
        def getSheetName(): String
        def toString(): String
    }

    pub enum SSTable(JTable)

    instance ToJava[SSTable] {
        type Out = JTable
        pub def toJava(x: SSTable): JTable = 
            let SSTable.SSTable(x1) = x;
            x1
    }

    instance ToFlix[SSTable] {
        type In = JTable
        pub def toFlix(x: JTable): SSTable = SSTable.SSTable(x)
    }

    pub def handle(table: SSTable, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - Table) = x -> 
        let jtable = ToJava.toJava(table);
        run {
            f(x)
        } with handler Table {

            def contains(cell, k) = {
                let cell1 = ToJava.toJava(cell);
                let ans = unsafe IO as {ef + Alt - Table} { jtable.contains(cell1) };
                k(ans)
            } 


            def getName(k) = {
                let ans = unsafe IO as {ef + Alt - Table} { jtable.getName() };
                k(ans)
            } 

            def getSheetName(k) = {
                let ans = unsafe IO as {ef + Alt - Table} { jtable.getSheetName() };
                k(ans)
            } 

            def toString(k) = {
                let ans = unsafe IO as {ef + Alt - Table} { jtable.toString() };
                k(ans)
            } 

        }


    pub def runCellReader(table: SSTable, f: Unit -> a \ ef): a \ (ef + Alt - Table) = 
        handle(table, f)()

    pub def runCellReaderWithResult(table: SSTable, f: Unit -> a \ ef): Result[String, a] \ (ef - Alt - Table) = 
        run { handle(table, f)() } with EffectExtras.Alt.runWithResult

}
