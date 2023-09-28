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


/// TODO - Use `Option[CellStyle]` as a Reader env in the Monad to allow inherited styles?

mod SheetIO {


    pub type alias RowWriter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.MutStateError[ka, Internal.Apache.Poi.SS.Usermodel.MutRow[r], ef, a]

}

mod SheetIO.RowWriter {

    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use GradedError.{swapError}
    use Graded.MutStateError.{liftSetter!, liftSetterResult!, liftGetter}

    use SheetIO.RowWriter
    use SheetIO.CellWriter

    use Internal.Apache.Poi.SS.Usermodel.MutCell
    use Internal.Apache.Poi.SS.Usermodel.CellStyle
    use Internal.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Internal.Apache.Poi.SS.Usermodel.MutRow
    use Internal.Apache.Poi.SS.Usermodel.MutRow.{MutRow}


    /// The state param `st` is expected to be mutated...
    pub def runRowWriter(ma: RowWriter[a, r, ef, a], 
                            st: MutRow[r]): Result[String, a] \ ef =
        Graded.MutStateError.runMutStateError(ma, st)

    // call `getOrCreateCell!` first ...
    pub def withCellWriter(write: CellWriter[a, r, ef, a], cell: MutCell[r]): RowWriter[ka, r, ef, a] =
        liftSetterResult!(_ -> SheetIO.CellWriter.runCellWriter(write, cell))


    def getCell(ix: Int32): RowWriter[ka, r, r, Option[MutCell[r]]] = 
        liftGetter(Internal.Apache.Poi.SS.Usermodel.MutRow.getCell(ix))

    def createCell!(ix: Int32): RowWriter[ka, r, r, MutCell[r]] = 
        liftSetterResult!(Internal.Apache.Poi.SS.Usermodel.MutRow.createCell!(ix))

    def getOrCreateCell!(ix: Int32): RowWriter[ka, r, r, MutCell[r]] = 
        getCell(ix) >>= ans -> 
        match ans {
            case Some(cell) => point(cell)
            case None       => createCell!(ix)
        }

    // call `getOrCreateCell!` first ...
    pub def writeAt!(ix: Int32, write: CellWriter[a, r, ef, a]): RowWriter[ka, r, ef + r, a] =
        getOrCreateCell!(ix)    >>= cell -> 
        withCellWriter(write, cell)


/*

    /// TODO should we set a style for the "get-or-create" cell as per the old implementation?

    /// This also sets the CellStyle...
    // def createOrGetCell(ix: Int32): RowWriter[r, ka, Cell[r]] \ r) = 
    //     use Result.{flatMap};
    //     use Internal.Apache.Poi.SS.Usermodel.Row.{createCell!};
    //     let body = (style, row) -> {
    //         let* cell   = match Internal.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
    //                         case Some(ans)  => if (not Object.isNull(ans)) Ok(ans) else createCell!(ix, row)
    //                         case None       => createCell!(ix, row)
    //                     };
    //         let _       = match style { 
    //                         case Some(sty1) => Internal.Apache.Poi.SS.Usermodel.Cell.setCellStyle!(sty1, cell)
    //                         case None       => unsafe_cast () as _ \ r
    //                     };
    //         Ok(cell)
    //     };
    //     liftSetterResult!(body)
*/


    pub def putHeaders!(v: Vector[String]): RowWriter[ka, r, r, Unit] = 
        def loop(ix) = match Vector.nth(ix, v) {
            case None    => point()
            case Some(s) => {
                writeAt!(ix, SheetIO.CellWriter.putString!(s)) >>= _ -> 
                loop(ix + 1)
            }
        };
        loop(0)


}

