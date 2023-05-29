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
    use Graded.Evaluator.{liftGetter, liftGetterResult, liftActionResult}

    use Basis.NIO.File.Path


    use Interop.Apache.Poi.SS.Usermodel.Sheet
    use Interop.Apache.Poi.SS.Usermodel.Sheet.{Sheet}
    use Interop.Apache.Poi.SS.Usermodel.Workbook
    use Interop.Apache.Poi.SS.Usermodel.Workbook.{Workbook}

    use SheetIO.WorkbookReader

    use SheetIO.SheetReader


    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: WorkbookReader[a, ef, a], path: Path): Result[String, a] \ {ef, IO} = 
        runXssf(ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(ma: WorkbookReader[a, ef, a], path: Path): Result[String, a] \ {ef, IO} = 
        runOnFile(ma, path) 

    /// Excel 97 binary (.xls) format.
    pub def runHssf(ma: WorkbookReader[a, ef, a], path: Path): Result[String, a] \ {ef, IO} = 
        runOnFile(ma, path)


    def runOnFile(ma: WorkbookReader[a, ef, a], path: Path): Result[String, a] \ {ef, IO} = region rc {
        forM (file  <- Basis.IO.FileInputStream.newWithPath(rc, path);
                ins <- Basis.IO.FileInputStream.toInputStream(file) |> Ok;
                wb  <- Interop.Apache.Poi.SS.Usermodel.WorkbookFactory.createWithInputStream!(ins);
                ans <- runWorkbookReader(ma, wb);
                _   <- Interop.Apache.Poi.SS.Usermodel.Workbook.close!(rc, wb))
            yield ans
    }




    pub def runWorkbookReader(ma: WorkbookReader[a, ef, a], 
                                st: Workbook): Result[String, a] \ ef =
        Graded.Evaluator.runEvaluator(ma, st)


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


    pub def getSheetNamed(name: String): WorkbookReader[ka, ef, Sheet] =
        liftGetterResult(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheet(name, wb)) |> Option.toOk("getSheetNamed")) 

    pub def getSheetAt(ix: Int32): WorkbookReader[ka, ef, Sheet] =
        liftGetterResult(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetAt(ix, wb)))


    pub def withSheetReader(eval: SheetReader[a, ef, a], sheet: Sheet): WorkbookReader[ka, ef, a] =
        liftActionResult(_ -> SheetIO.SheetReader.runSheetReader(eval, sheet))


}
