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
    use Basis.NIO.File.Path


    use Interop.Apache.Poi.SS.Usermodel.Sheet
    use Interop.Apache.Poi.SS.Usermodel.Sheet.{Sheet}
    use Interop.Apache.Poi.SS.Usermodel.Workbook
    use Interop.Apache.Poi.SS.Usermodel.Workbook.{Workbook}

    use SheetIO.WorkbookReader
    use SheetIO.RowEvaluator
    use SheetIO.NamedRowEvaluator
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator}

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


    pub def getWorkbookReader(x: WorkbookReader[r, ka, a]): Graded.Evaluator[r, ka, Workbook[r], a] = 
        let WorkbookReader(x1) = x;
        x1


    pub def runWorkbookReader(r: Region[r], 
                                ma: WorkbookReader[r, a, a], 
                                st: Workbook[r]): Result[a, String] \ r =
        Graded.Evaluator.runEvaluator(r, getWorkbookReader(ma), st)

    pub def point(x: a): WorkbookReader[r, ka, a] = 
        WorkbookReader(Graded.Evaluator.point(x))



    pub def map(f: a -> b \ ef, ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] \ ef =
        WorkbookReader(Graded.Evaluator.map(f, getWorkbookReader(ma)))

    pub def ap(mf: WorkbookReader[r, ka, a -> b \ ef], ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] \ ef =
        WorkbookReader(Graded.Evaluator.ap(getWorkbookReader(mf), getWorkbookReader(ma)))

    pub def flatMap(f: a -> WorkbookReader[r, ka, b] \ ef, ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] \ ef =
        WorkbookReader(Graded.Evaluator.flatMap(f >> getWorkbookReader, getWorkbookReader(ma)))

    pub def throwError(msg: String): WorkbookReader[r, ka, a] = 
        WorkbookReader(Graded.Evaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(Graded.Evaluator.mapError(f, getWorkbookReader(ma)))

    pub def catchError(ma: WorkbookReader[r, ka, a], f: String -> WorkbookReader[r, ka, a] \ ef): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(Graded.Evaluator.catchError(getWorkbookReader(ma), f >> getWorkbookReader))

    pub def empty(): WorkbookReader[r, ka, a] = 
        WorkbookReader(Graded.Evaluator.empty())
        
    pub def orelse(ma: WorkbookReader[r, ka, a], mb: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, a] =
        WorkbookReader(Graded.Evaluator.orelse(getWorkbookReader(ma), getWorkbookReader(mb)))

    def liftGetter(f: Workbook[r] -> a \ ef): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(Graded.Evaluator.liftGetter(f))
    
    def liftGetterResult(f: Workbook[r] -> Result[a, err] \ ef): WorkbookReader[r, ka, a] \ ef with ToString[err] =
        WorkbookReader(Graded.Evaluator.liftGetterResult(f))


    pub def liftAction(f: Unit -> a \ ef): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(Graded.Evaluator.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): WorkbookReader[r, ka, a] \ ef with ToString[err] =
        WorkbookReader(Graded.Evaluator.liftActionResult(f))



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

    pub def getNumberOfSheets(): WorkbookReader[r, ka, Int32] \ r =
        liftGetter(Interop.Apache.Poi.SS.Usermodel.Workbook.getNumberOfSheets)

    pub def getSheetNames(): WorkbookReader[r, ka, List[String]] \ r =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0,n);
        Traversable.traverse(getSheetNameWb, ixs)      


    def getSheetNameWb(ix: Int32): WorkbookReader[r, ka, String] \ r =
        liftGetter(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetName(ix))

    pub def getSheetNamed(name: String): WorkbookReader[r, ka, Sheet[r]] \ r =
        liftGetterResult(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheet(name) >> Option.toOk("getSheetNamed")) 

    pub def getSheetAt(ix: Int32): WorkbookReader[r, ka, Sheet[r]] \ r =
        liftGetterResult(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetAt(ix))

    pub def withSheetReader(eval: SheetReader[r, a, a], sheet: Sheet[r]): WorkbookReader[r, ka, a] \ r =
        liftActionResult(_ -> {let r = Scoped.regionOf(sheet); SheetIO.SheetReader.runSheetReader(r, eval, sheet)})

*/
}
