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

    use SheetIO/Internal/EvalError.{EvalError};
    
    pub enum WorkbookReader[r: Region, a: Type](
        ##org.apache.poi.ss.usermodel.Workbook -> Result[a, EvalError] & r
    )

    instance Functor[WorkbookReader[r]] {
        pub def map(f: a -> b & ef, o: WorkbookReader[r, a]): WorkbookReader[r, b] & ef = 
            SheetIO/WorkbookReader.map(f, o)
    }

    instance Applicative[WorkbookReader[r]] {
        pub def point(x: a): WorkbookReader[r, a] = 
            SheetIO/WorkbookReader.point(x)

        pub def ap(f: WorkbookReader[r, a -> b & ef], x: WorkbookReader[r, a]) : WorkbookReader[r, b] & ef = 
            SheetIO/WorkbookReader.ap(f, x)
    }

    instance Monad[WorkbookReader[r]] {
        pub def flatMap(f: a -> WorkbookReader[r, b] & ef, o: WorkbookReader[r, a]): WorkbookReader[r, b] & ef =
            SheetIO/WorkbookReader.flatMap(f, o)
    }

    instance MonadZero[WorkbookReader[r]] {
        pub def empty(): WorkbookReader[r, a] = SheetIO/WorkbookReader.throwError("empty")
    }

    instance MonadOr[WorkbookReader[r]] {
        pub def orelse(a: WorkbookReader[r, a], b: WorkbookReader[r, a]): WorkbookReader[r, a] =
            SheetIO/WorkbookReader.orelse(a, b)
    }


}

namespace SheetIO/WorkbookReader {

    use ToString.toString;
    use Interop/System.FilePath;

    use SheetIO.WorkbookReader;
    use SheetIO.WorkbookReader.{WorkbookReader};

    

    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO/Internal/PrimError.{PrimError};

    use SheetIO.Row;
    use SheetIO.Row.{Row};
    use SheetIO.RowEvaluator;
    use SheetIO.RowEvaluator.{RowEvaluator};
    use SheetIO.NamedRowEvaluator;
    use SheetIO.NamedRowEvaluator.{NamedRowEvaluator};
    use SheetIO.Sheet;
    use SheetIO.Sheet.{Sheet};
    use SheetIO.SheetReader;

    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: WorkbookReader[r, a], path: FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(ma: WorkbookReader[r, a], path: FilePath): Result[a, String] & Impure = 
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
    pub def runHssf(ma: WorkbookReader[r, a], path: FilePath): Result[a, String] & Impure = 
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


    def apply1(ma: WorkbookReader[r, a], wb: ##org.apache.poi.ss.usermodel.Workbook): Result[a, EvalError] & r = 
        let WorkbookReader(x1) = ma;
        x1(wb)

    pub def point(x: a): WorkbookReader[r, a] = 
        WorkbookReader(_ -> Ok(x) as & r)


    pub def throwError(msg: String): WorkbookReader[r, a] =
        WorkbookReader(_ -> 
            Err(SheetIO/Internal/EvalError.emptyEvalError({source = msg, exnMessage = ""})) as & r
        )

    pub def swapError(ma: WorkbookReader[r, a], msg: String): WorkbookReader[r, a] =
        WorkbookReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a)    => Ok(a)
                case Err(err) => Err({source = msg | err})
            }
        )


    pub def catchError(ma: WorkbookReader[r, a], catcher: EvalError -> WorkbookReader[r, a] & ef): WorkbookReader[r, a] & ef=
        WorkbookReader(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(e) => apply1(catcher(e) as & r, wb)
        }) as & ef


    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: WorkbookReader[r, a]): WorkbookReader[r, b] & ef = 
        WorkbookReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => { let ans = fn(a) as & r; Ok(ans) }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def ap(mf: WorkbookReader[r, a -> b & ef], ma: WorkbookReader[r, a]): WorkbookReader[r, b] & ef = 
        WorkbookReader(wb -> 
            match apply1(mf, wb) { 
                case Ok(f) => match apply1(ma, wb) { 
                    case Ok(a) => { let ans = f(a) as & r; Ok(ans) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        ) as & ef

    pub def flatMap(fn: a -> WorkbookReader[r, b] & ef, ma: WorkbookReader[r, a]): WorkbookReader[r, b] & ef = 
        WorkbookReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => apply1(fn(a) as & r, wb)
                case Err(e) => Err(e)
            }
        ) as & ef
  
    pub def bind(ma: WorkbookReader[r, a], fn: a -> WorkbookReader[r, b] & ef): WorkbookReader[r, b] & ef = 
        flatMap(fn, ma)


    pub def orelse(ma: WorkbookReader[r, a], mb: WorkbookReader[r, a]): WorkbookReader[r, a] =
        WorkbookReader(wb -> 
            match apply1(ma, wb) { 
                case Ok(a)  => Ok(a)
                case Err(_) => apply1(mb, wb)
            }
        )

    ///
    /// Expose the result of operation `ma`.
    ///
    /// By definition this never fails.
    ///
    pub def result(ma: WorkbookReader[r, a]): WorkbookReader[r, Result[a, String]] = 
        WorkbookReader(wb -> 
            let ans = apply1(ma, wb) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
            Ok(ans)
        )

    pub def guard(x: Bool): WorkbookReader[r, Unit] = 
        match x { 
            case true  => point()
            case false => throwError("guard")
        }


    def liftWorkbookPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Workbook -> Result[a, PrimError] & ef): WorkbookReader[r, a] & ef = 
        WorkbookReader(wb -> 
            let r = () as Region[r];
            {fn(r, wb) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.emptyEvalError)
        )  as & ef


    pub def liftAction(fn: Unit -> a & ef): WorkbookReader[r, a] & ef = 
        WorkbookReader(_ -> { let a = fn() as & r; Ok(a)}) as & ef


    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): WorkbookReader[r, a] & ef = 
        WorkbookReader(_ -> 
            {fn() as & r} |> Result.mapErr(s -> SheetIO/Internal/EvalError.emptyEvalError({ source = s, exnMessage = ""}))
        ) as & ef

    pub def liftOption(x: Option[a], errMsg: String): WorkbookReader[r, a] = match x { 
        case Some(x1) => point(x1)
        case None     => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): WorkbookReader[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): WorkbookReader[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): WorkbookReader[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }


    // SheetReader

    pub def getNumberOfSheets(): WorkbookReader[r, Int32] \ Read(r) =
        liftWorkbookPrim(SheetIO/Internal/Workbook.getNumberOfSheets)

    pub def getSheetNames(): WorkbookReader[r, List[String]] \ Read(r) =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0,n);
        Traversable.traverse(getSheetNameWb, ixs)      


    def getSheetNameWb(ix: Int32): WorkbookReader[r, String] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheetName(r, ix, wb))


    pub def getSheetNamed(name: String): WorkbookReader[r, Sheet[r]] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheet(r, name, wb) |> Result.map(Sheet)) 


    pub def getSheetAt(ix: Int32): WorkbookReader[r, Sheet[r]] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheetAt(r, ix, wb) |> Result.map(Sheet))

    pub def withSheetReader(eval: SheetReader[r, a], sheet: Sheet[r]): WorkbookReader[r, a] =
        WorkbookReader(_ -> SheetIO/SheetReader.evaluate(eval, sheet))


}
