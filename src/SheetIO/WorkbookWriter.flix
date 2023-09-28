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

    pub type alias WorkbookWriter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.MutStateError[ka, SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook[r], ef, a]

}

mod SheetIO.WorkbookWriter {

    use IOError.{Generic}
    use GradedMonad.{>>=}
    use Graded.MutStateError.{liftSetter!, liftSetterResult!, liftGetter, liftGetterResult}

    use Basis.IO.Path

    use SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook
    use SheetIO.Internal.Apache.Poi.SS.Usermodel.MutSheet
    use SheetIO.CellStyle
    use SheetIO.CellStyleSetter
    use SheetIO.DataFormat
    use SheetIO.Font
    use SheetIO.FontSetter
    use SheetIO.SheetWriter
    use SheetIO.WorkbookWriter


    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        runXssf(rc, ma, path)

    /// Excel 2007 OOXML (.xlsx) format.
    /// Path is for a *new* file...
    pub def runXssf(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        forM (fos       <- Basis.IO.FileOutputStream.new(path);
                wb      <- Ok(SheetIO.Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.new(rc));
                wb1     <- Ok(SheetIO.Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.asMutWorkbook(wb));
                ans     <- runWorkbookWriter(ma, wb1) |> Result.mapErr(Generic);
                outs    <- Ok(Basis.IO.FileOutputStream.asOutputStream(fos));
                poixml  <- Ok(SheetIO.Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.asPOIXMLDocument(wb));
                _       <- SheetIO.Internal.Apache.Poi.OOXML.POIXMLDocument.write!(outs, poixml) |> Result.mapErr(Generic);
                _       <- SheetIO.Internal.Apache.Poi.XSSF.Usermodel.XSSFWorkbook.close!(wb) |> Result.mapErr(Generic))
            yield ans


    /// Excel 2007 OOXML (.xlsx) format.
    pub def runAppendXlsx(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        runAppend(rc, ma, path)
    

    pub def runAppend(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        forM (fis       <- Basis.IO.FileInputStream.new(path);
                ins     <- Ok(Basis.IO.FileInputStream.asInputStream(fis));
                wb      <- SheetIO.Internal.Apache.Poi.SS.Usermodel.WorkbookFactory.createMutWithInputStream(rc, ins) |> Result.mapErr(Generic);
                ans     <- runWorkbookWriter(ma, wb) |> Result.mapErr(Generic);
                fos     <- Basis.IO.FileOutputStream.new(path);
                outs    <- Ok(Basis.IO.FileOutputStream.asOutputStream(fos));
                _       <- SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.write!(outs, wb) |> Result.mapErr(Generic);
                _       <- SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.close!(wb) |> Result.mapErr(Generic))
            yield ans


    /// Excel 2007 OOXML (.xlsx) format.
    /// Calls "newWithFile" not "new"...
    pub def runAppendXssf(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        runAppend(rc, ma, path)


    /// TODO - test...
    pub def runTemplateXssf(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], 
                            template: Path, 
                            outpath: Path): Result[IOError, a] \ {ef, IO, r} =
        let outpath1 = ToString.toString(outpath);
        forM (_         <- Files.copy(src = ToString.toString(template), outpath1);
                ans     <- runAppend(rc, ma, outpath))
            yield ans


    /// Excel 97 binary (.xls) format.
    pub def runHssf(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        forM (fos       <- Basis.IO.FileOutputStream.new(path);
                wb      <- Ok(SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.new(rc));
                wb1     <- Ok(SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.asMutWorkbook(wb));
                ans     <- runWorkbookWriter(ma, wb1) |> Result.mapErr(Generic);
                outs    <- Ok(Basis.IO.FileOutputStream.asOutputStream(fos));
                _       <- SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.write!(outs, wb) |> Result.mapErr(Generic);
                _       <- SheetIO.Internal.Apache.Poi.HSSF.Usermodel.HSSFWorkbook.close!(wb) |> Result.mapErr(Generic);
                _       <- Basis.IO.FileOutputStream.close!(fos))
            yield ans

    /// Excel 2007 OOXML (.xlsx) format.
    /// Calls "newWithFile" not "new"...
    pub def runAppendHssf(rc: Region[r], ma: WorkbookWriter[a, r, ef, a], path: Path): Result[IOError, a] \ {ef, IO, r} = 
        runAppend(rc, ma, path)


    pub def runTemplateHssf(rc: Region[r], 
                            ma: WorkbookWriter[a, r, ef, a], 
                            template: Path, 
                            outpath: Path): Result[IOError, a] \ {ef, IO, r} =
        let outpath1 = ToString.toString(outpath);
        forM (_         <- Files.copy(src = ToString.toString(template), outpath1);
                ans     <- runAppend(rc, ma, outpath))
            yield ans

    /// The state param `st` is expected to be mutated...
    pub def runWorkbookWriter(ma: WorkbookWriter[a, r, ef, a],  
                                st: MutWorkbook[r]): Result[String, a] \ ef =
        Graded.MutStateError.runMutStateError(ma, st) 


    pub def withSheetWriter(write: SheetWriter[a, r, ef, a], sheet: MutSheet[r]): WorkbookWriter[ka, r, ef, a] =
        Graded.MutStateError.liftGetterResult(_ -> SheetIO.SheetWriter.runSheetWriter(write, None, sheet))


    pub def createSheet!(name: String): WorkbookWriter[ka, r, r, MutSheet[r]] =
        liftSetterResult!(SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.createSheet!(name))

    pub def getSheetNamed(name: String): WorkbookWriter[ka, r, r, MutSheet[r]] =
        liftGetterResult(SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.getSheet(name) >> Option.toOk("getSheetNamed")) 


    pub def getSheetNames(): WorkbookWriter[ka, r, r, List[String]] =
        getNumberOfSheets() >>= n ->
        let ixs = List.range(0, n);
        GradedTraversable.traverse(getSheetName, ixs)      
    
    pub def getNumberOfSheets(): WorkbookWriter[ka, r, r, Int32] =
        liftGetter(SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.getNumberOfSheets)

    pub def getSheetName(ix: Int32): WorkbookWriter[ka, r, r, String] =
        liftGetter(SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.getSheetName(ix))

 
    pub def getSheetAt(ix: Int32): WorkbookWriter[ka, r, r, MutSheet[r]] =
        liftGetterResult(SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.getSheetAt(ix))

/*
    pub def getFormat!(format: String, df: DataFormat[r]): WorkbookWriter[r, ka, Int16] =
        liftAction(_ -> SheetIO.Internal.Apache.Poi.SS.Usermodel.DataFormat.getFormatNamed!(format, df))


    pub def createDataFormat!(): WorkbookWriter[r, ka, DataFormat[r]] =
        liftGetter(SheetIO.Internal.Apache.Poi.SS.Usermodel.MutWorkbook.createDataFormat!)


    pub def createFont!(mx: FontSetter[r, a, a]): WorkbookWriter[r, ka, Font[r]] \ r =
        let* font   = liftGetter(SheetIO.Internal.Apache.Poi.SS.Usermodel.Workbook.createFont!);
        let r1      = Scoped.regionOf(font);
        let* _      = liftActionResult(_ -> SheetIO.FontSetter.runFontSetter!(r1, mx, font));
        point(font)



    pub def createCellStyle!(rc: Region[r], mx: CellStyleSetter[r, a, a]): WorkbookWriter[r, ka, CellStyle[r]] \ r =
        let* style  = liftGetterResult(SheetIO.Internal.Apache.Poi.SS.Usermodel.Workbook.createCellStyle!);
        let* _      = liftActionResult(_ -> SheetIO.CellStyleSetter.runCellStyleSetter!(rc, mx, style));
        point(style)



    pub def withSheetWriter(writer: SheetWriter[r, a, a], sheet: Sheet[r]): WorkbookWriter[r, ka, a] \ r =
        let r1 = Scoped.regionOf(sheet);
        liftActionResult(_ -> SheetIO.SheetWriter.runSheetWriter!(r1, writer, None, sheet))




    // pub def liftAction(fn: Unit -> a \ IO): SheetWriter[r, a] = 
    //     SheetWriter((_, _, s1)-> {let a = fn(); Ok((s1, a)) })

    // pub def liftActionResult(fn: Unit -> Result[a, String] \ IO): SheetWriter[r, a] = 
    //     SheetWriter((_, _, s1) -> {let ans = fn(); Result.map(x -> (s1, x), ans)})


    // pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): SheetWriter[r, a] = match x { 
    //     case Ok(x1) => point(x1)
    //     case Err(e) => throwError(makeErrMsg(e))
    // }

    // pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetWriter[r, a] = match x { 
    //     case Success(x1) => point(x1)
    //     case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    // }

    // pub def mapWithIndexM(mf: (a, Int32) -> SheetWriter[r, b], xs: List[a]): SheetWriter[r, List[b]] =
    //     SheetWriter((wb, style, s1) ->
    //         mapWithIndexMHelper(wb, style, s1, mf, 0, xs, e -> Err(e), (s2, ys) -> Ok((s2, ys)))
    //     )

    // ///
    // /// Helper function for `mapM`.
    // ///
    // def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
    //                         style: Option[CellStyle],
    //                         s1: Cursors,
    //                         mf: (a, Int32) -> SheetWriter[r, b], 
    //                         ix : Int32, 
    //                         xs: List[a], 
    //                         fk: String -> Result[(Cursors, List[b]), String], 
    //                         sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), String]): Result[(Cursors, List[b]), String] \ IO =
    //     match xs {
    //         case Nil => sk(s1, Nil)
    //         case x :: rs => {
    //             let ma = mf(x, ix);
    //             match apply1(ma, wb, style, s1) {
    //                 case Err(e) => fk(e)
    //                 case Ok((s2, x1)) => mapWithIndexMHelper(wb, style, s2, mf, ix + 1, rs, fk, (s3, ks) -> sk(s3, x1 :: ks))
    //             }
    //         }
    //     }

    // pub def mapWithIndexX(mf: (a, Int32) -> SheetWriter[r, b], xs: List[a]): SheetWriter[r, Unit] =
    //     SheetWriter((wb, style, s1) ->
    //         mapWithIndexXHelper(wb, style, s1, mf, 0, xs, e -> Err(e), s2 -> Ok((s2, ())))
    //     )

    // def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
    //                         style: Option[CellStyle],
    //                         s1: Cursors,
    //                         mf: (a, Int32) -> SheetWriter[r, b],
    //                         ix: Int32, 
    //                         xs: List[a], 
    //                         fk: String -> Result[(Cursors, Unit), String], 
    //                         sk: Cursors -> Result[(Cursors, Unit), String]): Result[(Cursors, Unit), String] \ IO =
    //     match xs {
    //         case Nil => sk(s1)
    //         case x :: rs => {
    //             let ma = mf(x, ix);
    //             match apply1(ma, wb, style, s1) {
    //                 case Err(e) => fk(e)
    //                 case Ok((s2, _)) => mapWithIndexXHelper(wb, style, s2, mf, ix + 1, rs, fk, sk)
    //             }
    //         }
    //     }


    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    // pub def forM(xs: List[a], f: a -> SheetWriter[r, b]): SheetWriter[r, List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    // pub def forX(xs: List[a], f: a -> SheetWriter[r, b]): SheetWriter[r, Unit] = mapX(f, xs)


    // pub def foldM(mf: (b, a) -> SheetWriter[r, b], start: b, xs: List[a]): SheetWriter[r, b] =
    //     SheetWriter((wb, style, s1) ->
    //         foldMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac) -> Ok((s2, ac)))
    //     )

    // def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
    //                 style: Option[CellStyle],
    //                 s1: Cursors,
    //                 mf: (b, a) -> SheetWriter[r, b], 
    //                 acc: b, 
    //                 xs: List[a], 
    //                 fk: String -> Result[(Cursors, b), String], 
    //                 sk: (Cursors, b) -> Result[(Cursors, b), String]): Result[(Cursors, b), String] \ IO =
    //     match xs {
    //         case Nil => sk(s1, acc)
    //         case x :: rs => {
    //             let ma = mf(acc, x);
    //             match apply1(ma, wb, style, s1) {
    //                 case Err(e) => fk(e)
    //                 case Ok((s2, acc1)) => foldMHelper(wb, style, s2, mf, acc1, rs, fk, sk)
    //             }
    //         }
    //     }

    // pub def mapAccumM(mf: (ac, a) -> SheetWriter[r, (ac, b)], start: ac, xs: List[a]): SheetWriter[r, (ac, List[b])] =
    //     SheetWriter((wb, style, s1) ->
    //         mapAccumMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac, ks) -> Ok((s2, (ac, ks))))
    //     )

    // def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
    //                     style: Option[CellStyle],
    //                     s1: Cursors,
    //                     mf: (ac, a) -> SheetWriter[r, (ac, b)],
    //                     ac: ac,
    //                     xs: List[a],
    //                     fk: String -> Result[(Cursors, (ac, List[b])), String],
    //                     sk: (Cursors, ac, List[b]) -> Result[(Cursors, (ac, List[b])), String]): Result[(Cursors, (ac, List[b])), String] \ IO =
    //     match xs {
    //         case Nil => sk(s1, ac, Nil)
    //         case x :: rs => {
    //             let ma = mf(ac, x);
    //             match apply1(ma, wb, style, s1) {
    //                 case Err(e) => fk(e)
    //                 case Ok((s2, (ac1, x1))) => mapAccumMHelper(wb, style, s2, mf, ac1, rs, fk, (s3, ac2, ks) -> sk(s3, ac2, x1 :: ks))
    //             }
    //         }
    //     }

    // ///
    // /// `mapAccumX` is a synonym for `foldM`.
    // ///
    // pub def mapAccumX(mf: (ac, a) -> SheetWriter[r, ac], start: ac, xs: List[a]): SheetWriter[r, ac] = foldM(mf, start, xs)

    
    
    // /// Although this is `SheetWriter` some introspection (reading) is needed, 
    // /// e.g. if we are writing to a template file

*/

}
