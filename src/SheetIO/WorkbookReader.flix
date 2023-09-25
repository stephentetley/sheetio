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


    pub type alias WorkbookReader[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.RgnEnvEvaluator[ka, r, SheetIO.DataFormatter, Interop.Apache.Poi.SS.Usermodel.Workbook, ef, a]

}

mod SheetIO.WorkbookReader {

    use ToString.toString

    use GradedMonad.{>>=}
    use Graded.RgnEnvEvaluator.{liftGetter, liftGetterResult, liftActionResult}

    use Basis.IO.Path


    use Interop.Apache.Poi.SS.Usermodel.Sheet
    use Interop.Apache.Poi.SS.Usermodel.Sheet.{Sheet}
    use Interop.Apache.Poi.SS.Usermodel.Workbook
    use Interop.Apache.Poi.SS.Usermodel.Workbook.{Workbook}

    use SheetIO.WorkbookReader

    use SheetIO.SheetReader


    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(rc: Region[r], ma: WorkbookReader[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO} = 
        runXssf(rc, ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(rc: Region[r], ma: WorkbookReader[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO} = 
        runOnFile(rc, ma, path) 

    /// Excel 97 binary (.xls) format.
    pub def runHssf(rc: Region[r], ma: WorkbookReader[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO} = 
        runOnFile(rc, ma, path)


    def runOnFile(rc: Region[r], ma: WorkbookReader[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO} =
        forM (file  <- Basis.IO.FileInputStream.new(path);
                ins <- Ok(Basis.IO.FileInputStream.asInputStream(file));
                wb  <- Interop.Apache.Poi.SS.Usermodel.WorkbookFactory.createWithInputStream(ins);
                ans <- runWorkbookReader(ma, rc, wb);
                _   <- Interop.Apache.Poi.SS.Usermodel.Workbook.close!(rc, wb))
            yield ans




    pub def runWorkbookReader(ma: WorkbookReader[a, r, ef, a], 
                                rc: Region[r],
                                st: Workbook): Result[IOError, a] \ ef + r =        
        match Interop.Apache.Poi.SS.Usermodel.DataFormatter.new(rc, _ -> checked_ecast(Ok())){
            case Ok(df) => Graded.RgnEnvEvaluator.runRgnEnvEvaluator(ma, rc, df, st)
            case Err(msg) => Err(msg)
        }

    pub def withSheetReader(eval: SheetReader[a, r, ef, a], sheet: Sheet): WorkbookReader[ka, r, ef, a] =
        GradedExtras.Utils.liftGetterResultWithAll((rc, df, _) -> SheetIO.SheetReader.runSheetReader(eval, rc, df, sheet))

    pub def getSheetNamed(name: String): WorkbookReader[ka, r, ef, Sheet] =
        liftGetterResult(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheet(name, wb)) |> Option.toOk("getSheetNamed")) 

    
    pub def readSheetNamed(name: String, eval: SheetReader[a, r, ef, a]): WorkbookReader[ka, r, ef, a] =
        getSheetNamed(name)     >>= sheet -> 
        withSheetReader(eval, sheet)


    pub def getSheetAt(ix: Int32): WorkbookReader[ka, r, ef, Sheet] =
        liftGetterResult(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetAt(ix, wb)))
    
    pub def readSheetAt(ix: Int32, eval: SheetReader[a, r, ef, a]): WorkbookReader[ka, r, ef, a] =
        getSheetAt(ix)    >>= sheet -> 
        withSheetReader(eval, sheet)

    // WorkbookReader

    pub def getNumberOfSheets(): WorkbookReader[ka, r, ef, Int32] =
        liftGetter(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getNumberOfSheets(wb)))


    pub def getSheetNames(): WorkbookReader[ka, r, ef, List[String]] =
        getNumberOfSheets()     >>= n    -> {
            let ixs = List.range(0, n);
            GradedTraversable.traverse(getSheetName, ixs)
        }


    pub def getSheetName(ix: Int32): WorkbookReader[ka, r, ef, String] =
        liftGetter(wb -> checked_ecast(Interop.Apache.Poi.SS.Usermodel.Workbook.getSheetName(ix, wb)))





}
