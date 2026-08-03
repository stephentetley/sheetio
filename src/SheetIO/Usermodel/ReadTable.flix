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


pub mod SheetIO.Usermodel.ReadTable {
    
    
    
    use EffectExtras.Alt.Alt
    use SheetIO.Common.CellAddress
    use SheetIO.Common.CellType
    use SheetIO.Usermodel.{Table, Cell}

    import org.apache.poi.ss.usermodel.{Table => JTable}
    import java.lang.{Exception => JException}

    // Note table is only available to Xssf files (*.xlsx)

    pub eff ReadTable {
        def contains(cell: Cell): Bool
        def getName(): String
        def getSheetName(): String
        def toString(): String
    }



    pub def handle(table: Table, 
                    f: a -> b \ ef): a -> b \ (ef + Alt - ReadTable) = x -> 
        let jtable = ToJava.toJava(table);
        run {
            f(x)
        } with handler ReadTable {

            def contains(cell, k) = {
                let cell1 = ToJava.toJava(cell);
                let ans = unsafe IO as {ef + Alt - ReadTable} { jtable.contains(cell1) };
                k(ans)
            } 


            def getName(k) = {
                let ans = unsafe IO as {ef + Alt - ReadTable} { jtable.getName() };
                k(ans)
            } 

            def getSheetName(k) = {
                let ans = unsafe IO as {ef + Alt - ReadTable} { jtable.getSheetName() };
                k(ans)
            } 

            def toString(k) = {
                let ans = unsafe IO as {ef + Alt - ReadTable} { jtable.toString() };
                k(ans)
            } 

        }


    pub def runCellReader(table: Table, 
                            f: Unit -> a \ ef): a \ (ef + Alt - ReadTable) = 
        handle(table, f)()



}
