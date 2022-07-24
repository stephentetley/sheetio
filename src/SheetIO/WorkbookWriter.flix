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
    
    pub enum WorkbookWriter[r: Region, a: Type](
        ##org.apache.poi.ss.usermodel.Workbook -> Result[a, EvalError] & r)


    instance Functor[WorkbookWriter[r]] {
        pub def map(f: a -> b & ef, o: WorkbookWriter[r, a]): WorkbookWriter[r, b] & ef = 
            SheetIO/WorkbookWriter.map(f, o)
    }

    instance Applicative[WorkbookWriter[r]] {
        pub def point(x: a): WorkbookWriter[r, a] = 
            SheetIO/WorkbookWriter.point(x)

        pub def ap(f: WorkbookWriter[r, a -> b & ef], x: WorkbookWriter[r, a]) : WorkbookWriter[r, b] & ef = 
            SheetIO/WorkbookWriter.ap(f, x)
    }

    instance Monad[WorkbookWriter[r]] {
        pub def flatMap(f: a -> WorkbookWriter[r, b] & ef, o: WorkbookWriter[r, a]): WorkbookWriter[r, b] & ef =
            SheetIO/WorkbookWriter.flatMap(f, o)
    }


}

namespace SheetIO/WorkbookWriter {

    use SheetIO.WorkbookWriter;
    use SheetIO.WorkbookWriter.{WorkbookWriter};

    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.CellStyleSetter;
    use SheetIO.CellStyleSetter.{CellStyleSetter};
    use SheetIO.DataFormat;
    use SheetIO.DataFormat.{DataFormat};
    use SheetIO.Font;
    use SheetIO.Font.{Font};
    use SheetIO.FontSetter;
    use SheetIO.FontSetter.{FontSetter};    
    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO/Internal/PrimError.{PrimError};
    use SheetIO.Sheet;
    use SheetIO.Sheet.{Sheet};
    use SheetIO.SheetWriter;

    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: WorkbookWriter[r, a], 
                    path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)

    /// Excel 2007 OOXML (.xlsx) format.
    /// Path is for a *new* file...
    pub def runXssf(ma: WorkbookWriter[r, a], 
                    path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{newXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        region r {
            let outpath = ToString.toString(path);
            let* wb     = newXSSFWorkbook(r) |> Result.mapErr(SheetIO/Internal/PrimError.getText);
            let* ans    = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
            let* _      = writeXSSFWorkbook(wb, outpath);
            let* _      = closeXSSFWorkbook(wb);
            Ok(ans)
        }


    /// Excel 2007 OOXML (.xlsx) format.
    pub def runAppendXlsx(ma: WorkbookWriter[r, a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        runAppendXssf(ma, path)
    

    /// Excel 2007 OOXML (.xlsx) format.
    /// Calls "open" not "new"...
    pub def runAppendXssf(ma: WorkbookWriter[r, a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{openXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let path1 = ToString.toString(path);
        match openXSSFWorkbook(path1) { 
            case Err(e) => Err("runAppendXssf('${path1}'): ${e}")
            case Ok(wb) => {
                let* ans    = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
                let* _      = writeXSSFWorkbook(wb, path1);
                let* _      = closeXSSFWorkbook(wb);
                Ok(ans)
            }
        }


    /// TODO - is call to `writeXSSFWorkbook` the right thing to do?
    pub def runTemplateXssf(ma: WorkbookWriter[r, a], 
                            template: FlixSandbox/System.FilePath, 
                            outPath: FlixSandbox/System.FilePath): Result[a, String] & Impure =
        use SheetIO/Internal/Prim.{openXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let* _      = FlixSandbox/System/Directory.copy(template, outPath);
        let path1   = ToString.toString(outPath);
        let* wb     = openXSSFWorkbook(path1);
        let* ans    = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);        
        let* _      = writeXSSFWorkbook(wb, path1);
        let* _      = closeXSSFWorkbook(wb);
        Ok(ans)


    /// Excel 97 binary (.xls) format.
    pub def runHssf(ma: WorkbookWriter[r, a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{newHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};
        use Result.{flatMap};
        let path1   = ToString.toString(path);
        let* wb     = newHSSFWorkbook();
        let* ans    = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
        let* _      = writeHSSFWorkbook(wb, path1);
        let* _      = closeHSSFWorkbook(wb);
        Ok(ans)

    pub def runTemplateHssf(ma: WorkbookWriter[r, a], 
                            template: FlixSandbox/System.FilePath, 
                            outPath: FlixSandbox/System.FilePath): Result[a, String] & Impure =
        use Result.{flatMap};
        use SheetIO/Internal/Prim.{openHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};        
        let* _          = FlixSandbox/System/Directory.copy(template, outPath);
        let outPath1    = ToString.toString(outPath);
        let* wb         = openHSSFWorkbook(outPath1);
        let* ans        = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook) |> Result.mapErr(SheetIO/Internal/EvalError.getText);
        let* _          = writeHSSFWorkbook(wb, outPath1);
        let* _          = closeHSSFWorkbook(wb);
        Ok(ans)


    def apply1(ma: WorkbookWriter[r, a], 
               wb: ##org.apache.poi.ss.usermodel.Workbook): Result[a, EvalError] & r = 
        let WorkbookWriter(f1) = ma;
        f1(wb)

    pub def point(x: a): WorkbookWriter[r, a] = 
        WorkbookWriter(_ -> Ok(x) as & r)

    pub def throwError(msg: String): WorkbookWriter[r, a] =
        WorkbookWriter(_ -> 
            Err(SheetIO/Internal/EvalError.emptyEvalError({source = msg, exnMessage = ""})) as & r
        )

/*
    pub def catchError(ma: WorkbookWriter[r, a], handle: String -> WorkbookWriter[r, a]): WorkbookWriter[r, a] =
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => Ok((s2, a))
            case Err(e) => apply1(handle(e), wb, style, s1)
        })

    pub def mapError(f: String -> String, ma: SheetWriter[r, a]): SheetWriter[r, a] =
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => Ok((s2, a))
            case Err(e) => Err(f(e))
        })
*/

    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: WorkbookWriter[r, a]): WorkbookWriter[r, b] & ef = 
        WorkbookWriter(wb -> match apply1(ma, wb) { 
            case Ok(a) => { let a1 = fn(a) as & r; Ok(a1) }
            case Err(e) => Err(e)
        }) as & ef

    pub def ap(mf: WorkbookWriter[r, a -> b & ef], ma: WorkbookWriter[r, a]): WorkbookWriter[r, b] & ef = 
        WorkbookWriter(wb -> 
            match apply1(mf, wb) { 
                case Ok(f)  => match apply1(ma, wb) { 
                    case Ok(a)  => { let ans = f(a) as & r; Ok(ans) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        ) as & ef


    pub def flatMap(fn: a -> WorkbookWriter[r, b] & ef, ma: WorkbookWriter[r, a]): WorkbookWriter[r, b] & ef = 
        WorkbookWriter(wb -> match apply1(ma, wb) { 
            case Ok(a) => apply1(fn(a) as & r, wb)
            case Err(e) => Err(e)
        }) as & ef


    def liftWorkbookPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Workbook -> Result[a, PrimError] & ef): WorkbookWriter[r, a] & ef = 
        WorkbookWriter(wb -> 
            let r = () as Region[r];
            {fn(r, wb) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.emptyEvalError)
        )  as & ef

    pub def liftAction(fn: Unit -> a & ef): WorkbookWriter[r, a] & ef = 
        WorkbookWriter(_ -> { let a = fn() as & r; Ok(a)}) as & ef


    pub def liftActionResult(fn: Unit -> Result[a, String] & ef): WorkbookWriter[r, a] & ef = 
        WorkbookWriter(_ -> 
            {fn() as & r} |> Result.mapErr(s -> SheetIO/Internal/EvalError.emptyEvalError({ source = s, exnMessage = ""}))
        ) as & ef


    pub def createSheet(name: String): WorkbookWriter[r, Sheet[r]] \ Write(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.createSheet(r, name, wb) |> Result.map(Sheet))

    pub def getSheetNamed(name: String): WorkbookWriter[r, Sheet[r]] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheet(r, name, wb) |> Result.map(Sheet)) 


    pub def getSheetAt(ix: Int32): WorkbookWriter[r, Sheet[r]] \ Read(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.getSheetAt(r, ix, wb) |> Result.map(Sheet))

    pub def getFormat(x: DataFormat[r], format: String): WorkbookWriter[r, Int16] \ Read(r) =
        let DataFormat(x1) = x;
        liftWorkbookPrim((r, _) -> SheetIO/Internal/Util.getFormat(r, format, x1))


    pub def createDataFormat(): WorkbookWriter[r, DataFormat[r]] \ Write(r) =
        liftWorkbookPrim((r, wb) -> SheetIO/Internal/Workbook.createDataFormat(r, wb) |> Result.map(DataFormat))


    pub def createFont(mx: FontSetter[r, a]): WorkbookWriter[r, Font[r]] \ Write(r) =
        use SheetIO/Internal/Workbook.{createFont};
        let FontSetter(mf)      = mx;
        let* font               = liftWorkbookPrim(SheetIO/Internal/Workbook.createFont);
        let* _                  = WorkbookWriter(_ -> mf(font) |> Result.mapErr(SheetIO/Internal/EvalError.emptyEvalError));
        point(Font(font))


    pub def createCellStyle(mx: CellStyleSetter[r, a]): WorkbookWriter[r, CellStyle[r]] \ Write(r) =
        use SheetIO/Internal/Workbook.{createCellStyle};
        let CellStyleSetter(mf) = mx;
        let* style              = liftWorkbookPrim(SheetIO/Internal/Workbook.createCellStyle);
        let* _                  = WorkbookWriter(_ -> mf(style) |> Result.mapErr(SheetIO/Internal/EvalError.emptyEvalError));
        point(CellStyle(style))

    pub def withSheetWriter(writer: SheetWriter[r, a], sheet: Sheet[r]): WorkbookWriter[r, a] =
        WorkbookWriter(_ -> SheetIO/SheetWriter.write(writer, sheet))


/*


    pub def guard(x: Bool): SheetWriter[r, Unit] = 
        match x { 
            case true => point()
            case false => throwError("guard")
        }



    pub def liftAction(fn: Unit -> a & Impure): SheetWriter[r, a] = 
        SheetWriter((_, _, s1)-> {let a = fn(); Ok((s1, a)) })

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): SheetWriter[r, a] = 
        SheetWriter((_, _, s1) -> {let ans = fn(); Result.map(x -> (s1, x), ans)})

    pub def liftOption(x: Option[a], errMsg: String): SheetWriter[r, a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): SheetWriter[r, a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): SheetWriter[r, a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetWriter[r, a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }

    pub def localCellStyle(style: CellStyle, mx: SheetWriter[r, a]): SheetWriter[r, a] = 
        SheetWriter((wb, _, s1) -> apply1(mx, wb, Some(style), s1))




    pub def mapWithIndexM(mf: (a, Int32) -> SheetWriter[r, b], xs: List[a]): SheetWriter[r, List[b]] =
        SheetWriter((wb, style, s1) ->
            mapWithIndexMHelper(wb, style, s1, mf, 0, xs, e -> Err(e), (s2, ys) -> Ok((s2, ys)))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            style: Option[CellStyle],
                            s1: Cursors,
                            mf: (a, Int32) -> SheetWriter[r, b], 
                            ix : Int32, 
                            xs: List[a], 
                            fk: String -> Result[(Cursors, List[b]), String], 
                            sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), String]): Result[(Cursors, List[b]), String] & Impure =
        match xs {
            case Nil => sk(s1, Nil)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, x1)) => mapWithIndexMHelper(wb, style, s2, mf, ix + 1, rs, fk, (s3, ks) -> sk(s3, x1 :: ks))
                }
            }
        }

    pub def mapWithIndexX(mf: (a, Int32) -> SheetWriter[r, b], xs: List[a]): SheetWriter[r, Unit] =
        SheetWriter((wb, style, s1) ->
            mapWithIndexXHelper(wb, style, s1, mf, 0, xs, e -> Err(e), s2 -> Ok((s2, ())))
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            style: Option[CellStyle],
                            s1: Cursors,
                            mf: (a, Int32) -> SheetWriter[r, b],
                            ix: Int32, 
                            xs: List[a], 
                            fk: String -> Result[(Cursors, Unit), String], 
                            sk: Cursors -> Result[(Cursors, Unit), String]): Result[(Cursors, Unit), String] & Impure =
        match xs {
            case Nil => sk(s1)
            case x :: rs => {
                let ma = mf(x, ix);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, _)) => mapWithIndexXHelper(wb, style, s2, mf, ix + 1, rs, fk, sk)
                }
            }
        }


    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    pub def forM(xs: List[a], f: a -> SheetWriter[r, b]): SheetWriter[r, List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    pub def forX(xs: List[a], f: a -> SheetWriter[r, b]): SheetWriter[r, Unit] = mapX(f, xs)


    pub def foldM(mf: (b, a) -> SheetWriter[r, b], start: b, xs: List[a]): SheetWriter[r, b] =
        SheetWriter((wb, style, s1) ->
            foldMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac) -> Ok((s2, ac)))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    style: Option[CellStyle],
                    s1: Cursors,
                    mf: (b, a) -> SheetWriter[r, b], 
                    acc: b, 
                    xs: List[a], 
                    fk: String -> Result[(Cursors, b), String], 
                    sk: (Cursors, b) -> Result[(Cursors, b), String]): Result[(Cursors, b), String] & Impure =
        match xs {
            case Nil => sk(s1, acc)
            case x :: rs => {
                let ma = mf(acc, x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, acc1)) => foldMHelper(wb, style, s2, mf, acc1, rs, fk, sk)
                }
            }
        }

    pub def mapAccumM(mf: (ac, a) -> SheetWriter[r, (ac, b)], start: ac, xs: List[a]): SheetWriter[r, (ac, List[b])] =
        SheetWriter((wb, style, s1) ->
            mapAccumMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac, ks) -> Ok((s2, (ac, ks))))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        style: Option[CellStyle],
                        s1: Cursors,
                        mf: (ac, a) -> SheetWriter[r, (ac, b)],
                        ac: ac,
                        xs: List[a],
                        fk: String -> Result[(Cursors, (ac, List[b])), String],
                        sk: (Cursors, ac, List[b]) -> Result[(Cursors, (ac, List[b])), String]): Result[(Cursors, (ac, List[b])), String] & Impure =
        match xs {
            case Nil => sk(s1, ac, Nil)
            case x :: rs => {
                let ma = mf(ac, x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, (ac1, x1))) => mapAccumMHelper(wb, style, s2, mf, ac1, rs, fk, (s3, ac2, ks) -> sk(s3, ac2, x1 :: ks))
                }
            }
        }

    ///
    /// `mapAccumX` is a synonym for `foldM`.
    ///
    pub def mapAccumX(mf: (ac, a) -> SheetWriter[r, ac], start: ac, xs: List[a]): SheetWriter[r, ac] = foldM(mf, start, xs)



    // SheetWriter

    def getCursorAndIncr(sheetName: String): SheetWriter[r, Int32] = 
        SheetWriter((_, _, s1) -> match Map.get(sheetName, s1) {
            case Some(ix)   => Ok((Map.adjust(x -> x + 1, sheetName, s1), ix)) as & Impure
            case None       => Ok((Map.insert(sheetName, 1, s1), 0)) as & Impure
        })

    

    
    
    /// Although this is `SheetWriter` some introspection (reading) is needed, 
    /// e.g. if we are writing to a template file

    pub def getSheetNamed(name: String): SheetWriter[r, Sheet] =
        use SheetIO/Internal/Workbook.{getSheet};
        liftWorkbookPrim(x1 -> Result.map(Sheet, getSheet(x1, name))) 
    
    pub def getSheetNames(): SheetWriter[r, List[String]] =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0, n);
        mapM(getSheetNameWb, ixs)      
    
    def getNumberOfSheets(): SheetWriter[r, Int32] =
        use SheetIO/Internal/Workbook.{getNumberOfSheets};
        liftWorkbookPrim(getNumberOfSheets)

    def getSheetNameWb(ix: Int32): SheetWriter[r, String] =
        use SheetIO/Internal/Workbook.{getSheetName};
        liftWorkbookPrim(x1 -> getSheetName(x1, ix))


    pub def createDataFormat(): SheetWriter[r, DataFormat] =
        use SheetIO/Internal/Workbook.{createDataFormat};
        liftWorkbookPrim(x1 -> Result.map(DataFormat, createDataFormat(x1)))

    pub def getFormat(x: DataFormat, format: String): SheetWriter[r, Int16] =
        use SheetIO/Internal/Workbook.{getFormat};
        let DataFormat(x1) = x;
        liftActionResult(_ -> getFormat(x1, format))


    pub def setCursor!(x: Sheet, n: Int32): SheetWriter[r, Unit] =
        let* name = getSheetName(x);
        setCursorHelper(name, n)

    def getSheetName(x: Sheet): SheetWriter[r, String] =
        use SheetIO/Internal/Sheet.{getSheetName};
        let Sheet(x1) = x;
        liftActionResult(_ -> getSheetName(x1))

    def setCursorHelper(sheetName: String, n: Int32): SheetWriter[r, Unit] = 
        SheetWriter((_, _, s1) -> Ok((Map.insert(sheetName, n, s1), ())) as & Impure)


    pub def createCellStyle(mx: CellStyleSetter[r, a]): SheetWriter[r, CellStyle] =
        use SheetIO/Internal/Workbook.{createCellStyle};
        let CellStyleSetter(mf) = mx;
        let* style              = liftWorkbookPrim(createCellStyle >> Result.map(CellStyle));
        let CellStyle(style1)   = style;
        let* _                  = liftActionResult(_ -> mf(style1));
        point(style)



    pub def createFont(mx: FontSetter[r, a]): SheetWriter[r, Font] =
        use SheetIO/Internal/Workbook.{createFont};
        let FontSetter(mf)      = mx;
        let* font               = liftWorkbookPrim(createFont >> Result.map(Font));
        let Font(font1)         = font;
        let* _                  = liftActionResult(_ -> mf(font1));
        point(font)

    // createRow
    def createRow(x: Sheet, ix: Int32): SheetWriter[r, Row] =
        use SheetIO/Internal/Sheet.{createRow};
        let Sheet(x1) = x;
        liftActionResult(_ -> Result.map(Row, createRow(x1, ix)))

    pub def appendRow(x: Sheet, mx: RowSetter[a]): SheetWriter[r, Unit] =
        let RowSetter(mf) = mx;
        let* name       = getSheetName(x);
        let* ix         = getCursorAndIncr(name);
        let* row        = createRow(x, ix);
        let Row(row1)   = row;
        let* style      = askStyle();
        let* _          = liftActionResult(_ -> mf(row1, style));
        point()


    pub def allRows(sheet: Sheet, rows: List[a], writer: a -> RowSetter[Unit]): SheetWriter[r, Unit] = 
        mapX(r -> appendRow(sheet, writer(r)), rows)


    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allRowsWithHeaders(sheet: Sheet, headers: List[String], rows: List[a], writer: a -> RowSetter[Unit]): SheetWriter[r, Unit] = 
        let* _      = appendRow(sheet, SheetIO/RowSetter.putHeaders!(headers));
        let* _      = mapX(r -> appendRow(sheet, writer(r)), rows);
        point()

*/
}
