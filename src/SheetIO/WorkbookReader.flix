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


    pub type alias WorkbookReader[ka: Type, ef: Eff, a: Type] =
        Graded.Evaluator[ka, Interop.Apache.Poi.SS.Usermodel.Workbook, ef, a]

}

mod SheetIO.WorkbookReader {

    use ToString.toString

    use GradedMonad.{>>=}
    use Graded.Evaluator.{liftGetter}

    use Basis.NIO.File.Path


    use Interop.Apache.Poi.SS.Usermodel.Sheet
    use Interop.Apache.Poi.SS.Usermodel.Sheet.{Sheet}
    use Interop.Apache.Poi.SS.Usermodel.Workbook
    use Interop.Apache.Poi.SS.Usermodel.Workbook.{Workbook}

    use SheetIO.WorkbookReader

    use SheetIO.SheetReader

/*
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(r: Region[r], ma: WorkbookReader[r, a, a], path: Path): Result[a, String] \ IO = 
        runXssf(r, ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(r: Region[r], ma: WorkbookReader[r, a, a], path: Path): Result[a, String] \ IO = 
        runOnFile(r, ma, path) 

    def runOnFile(r: Region[r], ma: WorkbookReader[r, a, a], path: Path): Result[a, String] \ IO = 
        use Result.{flatMap};
        let path1   = toString(path);
        let file    = Interop.IO.File.newFile(r, path1);
        let* ins    = Interop.IO.FileInputStream.newWithFile(file);
        let* wb     = Interop.Apache.Poi.SS.Usermodel.WorkbookFactory.createWithInputStream!(ins);
        let* ans    = runWorkbookReader(r, ma, wb);
        let* _      = Interop.Apache.Poi.SS.Usermodel.Workbook.close!(wb);
        Ok(ans)
        

    /// Excel 97 binary (.xls) format.
    pub def runHssf(r: Region[r], ma: WorkbookReader[r, a, a], path: Path): Result[a, String] \ IO = 
        runOnFile(r, ma, path)

*/

    pub def runWorkbookReader(ma: WorkbookReader[a, ef, a], 
                                st: Workbook): Result[String, a] \ ef =
        Graded.Evaluator.runEvaluator(ma, st)



    ///
    /// Expose the result of operation `ma`.
    ///
    /// By definition this never fails.
    ///
    // pub def result(ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, Result[a, String]] = 
    //     WorkbookReader(wb -> 
    //         let ans = apply1(ma, wb) |> Result.mapErr(SheetIO.Internal.EvalError.getText);
    //         Ok(ans)
    //     )



    // WorkbookReader

    pub def getNumberOfSheets(): WorkbookReader[ka, ef, Int32] =
        liftGetter(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getNumberOfSheets(wb)))


    pub def getSheetNames(): WorkbookReader[ka, ef, List[String]] =
        getNumberOfSheets()     >>= n    -> {
            let ixs = List.range(0, n);
            GradedTraversable.traverse(getSheetName, ixs)
        }


    pub def getSheetName(ix: Int32): WorkbookReader[ka, ef, String] =
        liftGetter(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetName(ix, wb)))

/*
    pub def getSheetNamed(name: String): WorkbookReader[r, ka, Sheet[r]] \ r =
        liftGetterResult(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheet(name) >> Option.toOk("getSheetNamed")) 

    pub def getSheetAt(ix: Int32): WorkbookReader[r, ka, Sheet[r]] \ r =
        liftGetterResult(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetAt(ix))

    pub def withSheetReader(eval: SheetReader[r, a, a], sheet: Sheet[r]): WorkbookReader[r, ka, a] \ r =
        liftActionResult(_ -> {let r = Scoped.regionOf(sheet); SheetIO.SheetReader.runSheetReader(r, eval, sheet)})

*/
}
