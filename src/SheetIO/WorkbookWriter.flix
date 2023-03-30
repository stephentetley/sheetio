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



    pub enum WorkbookWriter[r: Region, ka: Type, a: Type](
        Graded.MutStateError[ka, Interop/Apache/Poi/SS/Usermodel.Workbook[r], r, a])

/*
    instance GradedFunctor[WorkbookWriter[r, ka]] {
        pub def map(f: a -> b & ef, o: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, b] & ef = 
            SheetIO/WorkbookWriter.map(f, o)
    }

    instance GradedApplicative[WorkbookWriter[r, ka]] {
        pub def point(x: a): WorkbookWriter[r, ka, a] = 
            SheetIO/WorkbookWriter.point(x)

        pub def ap(f: WorkbookWriter[r, ka, a -> b & ef], x: WorkbookWriter[r, ka, a]) : WorkbookWriter[r, ka, b] & ef = 
            SheetIO/WorkbookWriter.ap(f, x)
    }

    instance GradedMonad[WorkbookWriter[r, ka]] {
        pub def flatMap(f: a -> WorkbookWriter[r, ka, b] & ef, o: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, b] & ef =
            SheetIO/WorkbookWriter.flatMap(f, o)
    }

    instance GradedError[WorkbookWriter[r, ka]] {
        pub def throwError(msg: String): WorkbookWriter[r, ka, a] = 
            SheetIO/WorkbookWriter.throwError(msg)

        pub def mapError(f: String -> String \ ef, ma: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, a] \ ef =
            SheetIO/WorkbookWriter.mapError(f, ma)

        pub def catchError(ma: WorkbookWriter[r, ka, a], f: String -> WorkbookWriter[r, ka, a] \ ef): WorkbookWriter[r, ka, a] \ ef =
            SheetIO/WorkbookWriter.catchError(ma, f)
    }
*/

}

namespace SheetIO/WorkbookWriter {

    use Basis/NIO/File.Path;

    use SheetIO.WorkbookWriter;
    use SheetIO.WorkbookWriter.{WorkbookWriter};

    use Interop/Apache/Poi/SS/Usermodel.Workbook;
    use Interop/Apache/Poi/SS/Usermodel.Workbook.{Workbook};

    use SheetIO.CellStyle;
    use SheetIO.CellStyleSetter;
    use SheetIO.DataFormat;
    use SheetIO.Font;
    use SheetIO.FontSetter;
    use SheetIO.SheetWriter;
    
    use Interop/Apache/Poi/SS/Usermodel.Sheet;

/*
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(r: Region[r], ma: WorkbookWriter[r, a, a], path: Path): Result[a, String] \ {IO, Write(r)} = 
        runXssf(r, ma, path)

    /// Excel 2007 OOXML (.xlsx) format.
    /// Path is for a *new* file...
    pub def runXssf(r: Region[r], ma: WorkbookWriter[r, a, a], path: Path): Result[a, String] \ {IO, Write(r)} = 
        use Result.{flatMap};
        let outpath = ToString.toString(path);
        let* outs   = Interop/IO/FileOutputStream.newWithPath(outpath);
        let wb      = Interop/Apache/Poi/XSSF/Usermodel/XSSFWorkbook.new(r);
        let wb1     = Interop/Apache/Poi/SS/Usermodel/SubWorkbook.toWorkbook(wb);
        let* ans    = runWorkbookWriter!(r, ma, wb1);
        let* _      = Interop/Apache/Poi/OOXML/POIXMLDocument.write!(outs, wb);
        let* _      = Interop/Apache/Poi/XSSF/Usermodel/XSSFWorkbook.close!(wb);
        Ok(ans)



    /// Excel 2007 OOXML (.xlsx) format.
    pub def runAppendXlsx(r: Region[r], ma: WorkbookWriter[r, a, a], path: Path): Result[a, String] \ IO = 
        runAppendXssf(r, ma, path)
    

    /// Excel 2007 OOXML (.xlsx) format.
    /// Calls "newWithFile" not "new"...
    pub def runAppendXssf(r: Region[r], ma: WorkbookWriter[r, a, a], path: Path): Result[a, String] \ IO = 
        use Result.{flatMap};
        let path1   = ToString.toString(path);
        let file    = Interop/IO/File.newFile(r, path1);
        let* wb     = Interop/Apache/Poi/XSSF/Usermodel/XSSFWorkbook.newWithFile(file);
        let wb1     = Interop/Apache/Poi/SS/Usermodel/SubWorkbook.toWorkbook(wb);
        let* ans    = runWorkbookWriter!(r, ma, wb1);
        let* outs   = Interop/IO/FileOutputStream.newWithPath(path1);
        let* _      = Interop/Apache/Poi/OOXML/POIXMLDocument.write!(outs, wb);
        let* _      = Interop/Apache/Poi/XSSF/Usermodel/XSSFWorkbook.close!(wb);
        Ok(ans)


    /// TODO - is call to `writeXSSFWorkbook` the right thing to do?
    // pub def runTemplateXssf(ma: WorkbookWriter[r, a, a], 
    //                         template: Path, 
    //                         outPath: Path): Result[a, String] & Impure =
    //     use SheetIO/Internal/Prim.{openXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
    //     use Result.{flatMap};
    //     let* _      = File.copy(src = ToString.toString(template), ToString.toString(outPath));
    //     let path1   = ToString.toString(outPath);
    //     let* wb     = openXSSFWorkbook(path1) as ##org.apache.poi.ss.usermodel.Workbook;
    //     let* ans    = runWorkbookWriter!(() as Region[r], ma, wb);
    //     let* _      = writeXSSFWorkbook(wb, path1);
    //     let* _      = closeXSSFWorkbook(wb);
    //     Ok(ans)


    /// Excel 97 binary (.xls) format.
    // pub def runHssf(ma: WorkbookWriter[r, a, a], path: Path): Result[a, String] & Impure = 
    //     use SheetIO/Internal/Prim.{newHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};
    //     use Result.{flatMap};
    //     let path1   = ToString.toString(path);
    //     let* wb     = newHSSFWorkbook() as ##org.apache.poi.ss.usermodel.Workbook;
    //     let* ans    = runWorkbookWriter!(() as Region[r], ma, wb);
    //     let* _      = writeHSSFWorkbook(wb, path1);
    //     let* _      = closeHSSFWorkbook(wb);
    //     Ok(ans)

    // pub def runTemplateHssf(ma: WorkbookWriter[r, a, a], 
    //                         template: Path, 
    //                         outPath: Path): Result[a, String] & Impure =
    //     use Result.{flatMap};
    //     use SheetIO/Internal/Prim.{openHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};        
    //     let* _          = File.copy(src = ToString.toString(template), ToString.toString(outPath));
    //     let outPath1    = ToString.toString(outPath);
    //     let* wb         = openHSSFWorkbook(outPath1) as ##org.apache.poi.ss.usermodel.Workbook;
    //     let* ans        = runWorkbookWriter!(() as Region[r], ma, wb);
    //     let* _          = writeHSSFWorkbook(wb, outPath1);
    //     let* _          = closeHSSFWorkbook(wb);
    //     Ok(ans)

    pub def getWorkbookWriter(x: WorkbookWriter[r, ka, a]): Graded.MutStateError[r, ka, Workbook[r], a] = 
        let WorkbookWriter(x1) = x;
        x1

    /// The state param `st` is expected to be mutated...
    pub def runWorkbookWriter!(r: Region[r], 
                                ma: WorkbookWriter[r, a, a], 
                                st: Workbook[r]): Result[a, String] \ r =
        Graded/MutStateError.runMutStateError(r, getWorkbookWriter(ma), st)

    pub def point(x: a): WorkbookWriter[r, ka, a] = 
        WorkbookWriter(Graded/MutStateError.point(x))


    pub def map(f: a -> b \ ef, x: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, b] \ ef =
        WorkbookWriter(Graded/MutStateError.map(f, getWorkbookWriter(x)))

    pub def ap(mf: WorkbookWriter[r, ka, a -> b \ ef], ma: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, b] \ ef =
        WorkbookWriter(Graded/MutStateError.ap(getWorkbookWriter(mf), getWorkbookWriter(ma)))

    pub def flatMap(k: a -> WorkbookWriter[r, ka, b] \ ef, x: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, b]  \ ef = 
        WorkbookWriter(Graded/MutStateError.flatMap(k >> getWorkbookWriter, getWorkbookWriter(x)))


    pub def throwError(msg: String): WorkbookWriter[r, ka, a] = 
        WorkbookWriter(Graded/MutStateError.throwError(msg))

    pub def mapError(f: String -> String \ ef, ma: WorkbookWriter[r, ka, a]): WorkbookWriter[r, ka, a] \ ef =
        WorkbookWriter(Graded/MutStateError.mapError(f, getWorkbookWriter(ma)))

    pub def catchError(ma: WorkbookWriter[r, ka, a], f: String -> WorkbookWriter[r, ka, a] \ ef): WorkbookWriter[r, ka, a] \ ef =
        WorkbookWriter(Graded/MutStateError.catchError(getWorkbookWriter(ma), f >> getWorkbookWriter))

    // def liftSetter!(f: Workbook[r] -> a \ ef): WorkbookWriter[r, ka, a] \ ef  =
    //     WorkbookWriter(Graded/MutStateError.liftSetter!(f))

    def liftSetterResult!(f: Workbook[r] -> Result[a, err] \ ef): WorkbookWriter[r, ka, a] \ ef with ToString[err] =
        WorkbookWriter(Graded/MutStateError.liftSetterResult!(f))

    def liftGetter(f: Workbook[r] -> a \ ef): WorkbookWriter[r, ka, a] \ ef  =
        WorkbookWriter(Graded/MutStateError.liftGetter(f))

    def liftGetterResult(f: Workbook[r] -> Result[a, err] \ ef): WorkbookWriter[r, ka, a] \ ef with ToString[err] =
        WorkbookWriter(Graded/MutStateError.liftGetterResult(f))


    pub def liftAction(f: Unit -> a \ ef): WorkbookWriter[r, ka, a] \ ef =
        WorkbookWriter(Graded/MutStateError.liftAction(f))
    
        
    pub def liftActionResult(f: Unit -> Result[a, err] \ ef): WorkbookWriter[r, ka, a] \ ef with ToString[err] =
        WorkbookWriter(Graded/MutStateError.liftActionResult(f))



    pub def createSheet!(name: String): WorkbookWriter[r, ka, Sheet[r]] \ Write(r) =
        liftSetterResult!(Interop/Apache/Poi/SS/Usermodel/Workbook.createSheet!(name))


    pub def getSheetNamed(name: String): WorkbookWriter[r, ka, Sheet[r]] \ Read(r) =
        liftGetterResult(Interop/Apache/Poi/SS/Usermodel/Workbook.getSheet(name) >> Option.toOk("getSheetNamed")) 


    pub def getSheetNames(): WorkbookWriter[r, ka, List[String]] \ Read(r) =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0, n);
        Traversable.traverse(getSheetName, ixs)      
    
    pub def getNumberOfSheets(): WorkbookWriter[r, ka, Int32] \ Read(r) =
        liftGetter(Interop/Apache/Poi/SS/Usermodel/Workbook.getNumberOfSheets)

    pub def getSheetName(ix: Int32): WorkbookWriter[r, ka, String] \ Read(r) =
        liftGetter(Interop/Apache/Poi/SS/Usermodel/Workbook.getSheetName(ix))

 
    pub def getSheetAt(ix: Int32): WorkbookWriter[r, ka, Sheet[r]] \ Read(r) =
        liftGetterResult(Interop/Apache/Poi/SS/Usermodel/Workbook.getSheetAt(ix))


    pub def getFormat!(format: String, df: DataFormat[r]): WorkbookWriter[r, ka, Int16] \ Write(r) =
        liftAction(_ -> Interop/Apache/Poi/SS/Usermodel/DataFormat.getFormatNamed!(format, df))


    pub def createDataFormat!(): WorkbookWriter[r, ka, DataFormat[r]] \ Write(r) =
        liftGetter(Interop/Apache/Poi/SS/Usermodel/Workbook.createDataFormat!)


    pub def createFont!(mx: FontSetter[r, a, a]): WorkbookWriter[r, ka, Font[r]] \ Write(r) =
        let* font   = liftGetter(Interop/Apache/Poi/SS/Usermodel/Workbook.createFont!);
        let r1      = Scoped.regionOf(font);
        let* _      = liftActionResult(_ -> SheetIO/FontSetter.runFontSetter!(r1, mx, font));
        point(font)



    pub def createCellStyle!(rc: Region[r], mx: CellStyleSetter[r, a, a]): WorkbookWriter[r, ka, CellStyle[r]] \ Write(r) =
        let* style  = liftGetterResult(Interop/Apache/Poi/SS/Usermodel/Workbook.createCellStyle!);
        let* _      = liftActionResult(_ -> SheetIO/CellStyleSetter.runCellStyleSetter!(rc, mx, style));
        point(style)



    pub def withSheetWriter(writer: SheetWriter[r, a, a], sheet: Sheet[r]): WorkbookWriter[r, ka, a] \ Write(r) =
        let r1 = Scoped.regionOf(sheet);
        liftActionResult(_ -> SheetIO/SheetWriter.runSheetWriter!(r1, writer, None, sheet))




    // pub def liftAction(fn: Unit -> a & Impure): SheetWriter[r, a] = 
    //     SheetWriter((_, _, s1)-> {let a = fn(); Ok((s1, a)) })

    // pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): SheetWriter[r, a] = 
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
    //                         sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), String]): Result[(Cursors, List[b]), String] & Impure =
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
    //                         sk: Cursors -> Result[(Cursors, Unit), String]): Result[(Cursors, Unit), String] & Impure =
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
    //                 sk: (Cursors, b) -> Result[(Cursors, b), String]): Result[(Cursors, b), String] & Impure =
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
    //                     sk: (Cursors, ac, List[b]) -> Result[(Cursors, (ac, List[b])), String]): Result[(Cursors, (ac, List[b])), String] & Impure =
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
