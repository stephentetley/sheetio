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


namespace SheetIO {


    pub enum WorkbookReader[r: Region, ka: Type, a: Type](
        MonadLib.Evaluator[r, ka, Interop/Apache/Poi/SS/Usermodel.Workbook[r], a])


    instance Functor[WorkbookReader[r, ka]] {
        pub def map(f: a -> b & ef, o: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] & ef = 
            SheetIO/WorkbookReader.map(f, o)
    }

    instance Applicative[WorkbookReader[r, ka]] {
        pub def point(x: a): WorkbookReader[r, ka, a] = 
            SheetIO/WorkbookReader.point(x)

        pub def ap(f: WorkbookReader[r, ka, a -> b & ef], x: WorkbookReader[r, ka, a]) : WorkbookReader[r, ka, b] & ef = 
            SheetIO/WorkbookReader.ap(f, x)
    }

    instance Monad[WorkbookReader[r, ka]] {
        pub def flatMap(f: a -> WorkbookReader[r, ka, b] & ef, o: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] & ef =
            SheetIO/WorkbookReader.flatMap(f, o)
    }

    instance MonadZero[WorkbookReader[r, ka]] {
        pub def empty(): WorkbookReader[r, ka, a] = 
            SheetIO/WorkbookReader.empty()
    }

    instance MonadLib.MonadOr[WorkbookReader[r, ka]] {
        pub def orelse(a: WorkbookReader[r, ka, a], b: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, a] =
            SheetIO/WorkbookReader.orelse(a, b)
    }


}

namespace SheetIO/WorkbookReader {

    use ToString.toString;
    use Basis/System.FilePath;

    use SheetIO.WorkbookReader;
    use SheetIO.WorkbookReader.{WorkbookReader};

    use Interop/Apache/Poi/SS/Usermodel.Sheet;
    use Interop/Apache/Poi/SS/Usermodel.Sheet.{Sheet};
    use Interop/Apache/Poi/SS/Usermodel.Workbook;
    use Interop/Apache/Poi/SS/Usermodel.Workbook.{Workbook};

    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};

    use SheetIO.SheetReader;

/*
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: WorkbookReader[r, ka, a], path: FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(ma: WorkbookReader[r, ka, a], path: FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{openXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let path1 = toString(path);
        match openXSSFWorkbook(path1) { 
            case Err(e) => Err("runXssf(${path1}): ${e}")
            case Ok(wb) => {
                let ans = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
                let* _ = closeXSSFWorkbook(wb);
                ans         
            }
        }

    /// Excel 97 binary (.xls) format.
    pub def runHssf(ma: WorkbookReader[r, ka, a], path: FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{openHSSFWorkbook, closeHSSFWorkbook};
        use Result.{flatMap};
        let path1 = toString(path);
        match openHSSFWorkbook(path1) { 
            case Err(e) => Err("runHssf(${path1}): ${e}")
            case Ok(wb) => {
                let ans = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
                let* _ = closeHSSFWorkbook(wb);
                ans         
            }
        }
*/


    pub def getWorkbookReader(x: WorkbookReader[r, ka, a]): MonadLib.Evaluator[r, ka, Workbook[r], a] = 
        let WorkbookReader(x1) = x;
        x1

    /// The state param `st` is expected to be mutated...
    pub def runWorkbookReader(r: Region[r], 
                                ma: WorkbookReader[r, a, a], 
                                st: Workbook[r]): Result[a, String] \ r =
        MonadLib/Evaluator.runEvaluator(r, getWorkbookReader(ma), st)

    pub def point(x: a): WorkbookReader[r, ka, a] = 
        WorkbookReader(MonadLib/Evaluator.point(x))



    pub def map(f: a -> b \ ef, ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] \ ef =
        WorkbookReader(MonadLib/Evaluator.map(f, getWorkbookReader(ma)))

    pub def ap(mf: WorkbookReader[r, ka, a -> b \ ef], ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] \ ef =
        WorkbookReader(MonadLib/Evaluator.ap(getWorkbookReader(mf), getWorkbookReader(ma)))

    pub def flatMap(f: a -> WorkbookReader[r, ka, b] \ ef, ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, b] \ ef =
        WorkbookReader(MonadLib/Evaluator.flatMap(f >> getWorkbookReader, getWorkbookReader(ma)))

    pub def throwError(msg: String): WorkbookReader[r, ka, a] = 
        WorkbookReader(MonadLib/Evaluator.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(MonadLib/Evaluator.mapError(f, getWorkbookReader(ma)))

    pub def catchError(ma: WorkbookReader[r, ka, a], f: String -> WorkbookReader[r, ka, a] \ ef): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(MonadLib/Evaluator.catchError(getWorkbookReader(ma), f >> getWorkbookReader))

    pub def empty(): WorkbookReader[r, ka, a] = 
        WorkbookReader(MonadLib/Evaluator.empty())
        
    pub def orelse(ma: WorkbookReader[r, ka, a], mb: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, a] =
        WorkbookReader(MonadLib/Evaluator.orelse(getWorkbookReader(ma), getWorkbookReader(mb)))

    def _liftGetter(f: Workbook[r] -> a \ ef): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(MonadLib/Evaluator.liftGetter(f))
    
    def _liftGetterResult(f: Workbook[r] -> Result[a, err] \ ef): WorkbookReader[r, ka, a] \ ef with ToString[err] =
        WorkbookReader(MonadLib/Evaluator.liftGetterResult(f))


    pub def liftAction(f: Unit -> a \ ef): WorkbookReader[r, ka, a] \ ef =
        WorkbookReader(MonadLib/Evaluator.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): WorkbookReader[r, ka, a] \ ef with ToString[err] =
        WorkbookReader(MonadLib/Evaluator.liftActionResult(f))


/*
    ///
    /// Expose the result of operation `ma`.
    ///
    /// By definition this never fails.
    ///
    pub def result(ma: WorkbookReader[r, ka, a]): WorkbookReader[r, ka, Result[a, String]] = 
        WorkbookReader(wb -> 
            let ans = apply1(ma, wb) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
            Ok(ans)
        )

    pub def guard(x: Bool): WorkbookReader[r, ka, Unit] = 
        match x { 
            case true  => point()
            case false => throwError("guard")
        }
*/


/*
    pub def liftOption(x: Option[a], errMsg: String): WorkbookReader[r, ka, a] = match x { 
        case Some(x1) => point(x1)
        case None     => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): WorkbookReader[r, ka, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): WorkbookReader[r, ka, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): WorkbookReader[r, ka, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }
*/

/*

    // WorkbookReader

    pub def getNumberOfSheets(): WorkbookReader[r, ka, Int32] \ Read(r) =
        liftWorkbookPrim(SheetIO/Internal/Workbook.getNumberOfSheets)

    pub def getSheetNames(): WorkbookReader[r, ka, List[String]] \ Read(r) =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0,n);
        Traversable.traverse(getSheetNameWb, ixs)      


    def getSheetNameWb(ix: Int32): WorkbookReader[r, ka, String] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheetName(r, ix, wb))


    pub def getSheetNamed(name: String): WorkbookReader[r, ka, Sheet[r]] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheet(r, name, wb) |> Result.map(Sheet)) 


    pub def getSheetAt(ix: Int32): WorkbookReader[r, ka, Sheet[r]] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheetAt(r, ix, wb) |> Result.map(Sheet))
*/

    pub def withSheetReader(eval: SheetReader[r, a, a], sheet: Sheet[r]): WorkbookReader[r, ka, a] \ Read(r) =
        liftActionResult(_ -> {let r = () as Region[r]; SheetIO/SheetReader.runSheetReader(r, eval, sheet)})


}
