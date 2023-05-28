/*
 * Copyright 2023 Stephen Tetley
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

    pub type alias CellReader[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.EnvEvaluator[ka, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Cell[r], ef, a]

    
}

mod SheetIO.CellReader {

    use GradedFunctor.{<$$>}
    use Graded.EnvEvaluator.{liftGetter, liftGetterWithEnv}

    use SheetIO.CellReader
    use SheetIO.CellType
    use SheetIO.CellType.{TyNone, TyBlank, TyBoolean, TyError, TyFormula, TyNumeric, TyString}

    /// Uses the default DataFormatter
    pub def getContent(): CellReader[ka, r, r, String] =
        liftGetterWithEnv((fmt, cell) -> Interop.Apache.Poi.SS.Usermodel.DataFormatter.formatCellValue(cell, fmt))

    /// Uses the default DataFormatter
    pub def getTrimmedContent(): CellReader[ka, r, r, String] =
        getContent() <$$> String.trim

    pub def isNull(): CellReader[ka, r, r, Bool] = 
        liftGetter(cell -> checked_ecast(Object.isNull(cell)))

    pub def getCellType(): CellReader[ka, r, r, CellType] =
        liftGetter(cell -> 
            Interop.Apache.Poi.SS.Usermodel.Cell.getCellType(cell) 
                |> SheetIO.CellType.fromNative
        )


}