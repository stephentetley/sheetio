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

    use SheetIO.CellStyle;
   
    type alias Cursors = Map[String, Int32]
    
    pub enum SheetWriter[a](##org.apache.poi.ss.usermodel.Workbook -> Option[CellStyle] -> Cursors -> Result[(Cursors, a), String] & Impure)


    instance Functor[SheetWriter] {
        pub def map(f: a -> b & ef, o: SheetWriter[a]): SheetWriter[b] & ef = 
            SheetIO/SheetWriter.map(f, o)
    }

    instance Applicative[SheetWriter] {
        pub def point(x: a): SheetWriter[a] = 
            SheetIO/SheetWriter.point(x)

        pub def ap(f: SheetWriter[a -> b & ef], x: SheetWriter[a]) : SheetWriter[b] & ef = 
            SheetIO/SheetWriter.ap(f, x)
    }

    instance Monad[SheetWriter] {
        pub def flatMap(f: a -> SheetWriter[b] & ef, o: SheetWriter[a]): SheetWriter[b] & ef =
            SheetIO/SheetWriter.flatMap(f, o)
    }


}

namespace SheetIO/SheetWriter {

    use ToString.toString;

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
    use SheetIO.Row;
    use SheetIO.Row.{Row};
    use SheetIO.RowSetter;
    use SheetIO.RowSetter.{RowSetter};
    use SheetIO.Sheet;
    use SheetIO.Sheet.{Sheet};

    use SheetIO.SheetWriter;
    use SheetIO.SheetWriter.{SheetWriter};
    use SheetIO.Cursors;

    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(ma: SheetWriter[a], 
                    path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        runXssf(ma, path)

    /// Excel 2007 OOXML (.xlsx) format.
    /// Path is for a *new* file...
    pub def runXssf(ma: SheetWriter[a], 
                    path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{newXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let outpath     = toString(path);
        let* wb         = newXSSFWorkbook();
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
        let* _          = writeXSSFWorkbook(wb, outpath);
        let* _          = closeXSSFWorkbook(wb);
        Ok(ans)

    /// Excel 2007 OOXML (.xlsx) format.
    pub def runAppendXlsx(ma: SheetWriter[a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        runAppendXssf(ma, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runAppendXssf(ma: SheetWriter[a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{openXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let path1 = toString(path);
        match openXSSFWorkbook(path1) { 
            case Err(e) => Err("runAppendXssf('${path1}'): ${e}")
            case Ok(wb) => {
                let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
                let* _          = writeXSSFWorkbook(wb, path1);
                let* _          = closeXSSFWorkbook(wb);
                Ok(ans)
            }
        }

    /// TODO - is call to `writeXSSFWorkbook` the right thing to do?
    pub def runTemplateXssf(ma: SheetWriter[a], 
                            template: FlixSandbox/System.FilePath, 
                            outPath: FlixSandbox/System.FilePath): Result[a, String] & Impure =
        use SheetIO/Internal/Prim.{openXSSFWorkbook, writeXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let* _          = FlixSandbox/System/Directory.copy(template, outPath);
        let path1       = toString(outPath);
        let* wb         = openXSSFWorkbook(path1);
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());        
        let* _          = writeXSSFWorkbook(wb, path1);
        let* _          = closeXSSFWorkbook(wb);
        Ok(ans)

                    
    /// Excel 97 binary (.xls) format.
    pub def runHssf(ma: SheetWriter[a], path: FlixSandbox/System.FilePath): Result[a, String] & Impure = 
        use SheetIO/Internal/Prim.{newHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};
        use Result.{flatMap};
        let path1       = toString(path);
        let* wb         = newHSSFWorkbook();
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
        let* _          = writeHSSFWorkbook(wb, path1);
        let* _          = closeHSSFWorkbook(wb);
        Ok(ans)

    pub def runTemplateHssf(ma: SheetWriter[a], 
                            template: FlixSandbox/System.FilePath, 
                            outPath: FlixSandbox/System.FilePath): Result[a, String] & Impure =
        use Result.{flatMap};
        use SheetIO/Internal/Prim.{openHSSFWorkbook, writeHSSFWorkbook, closeHSSFWorkbook};        
        let* _          = FlixSandbox/System/Directory.copy(template, outPath);
        let outPath1    = toString(outPath);
        let* wb         = openHSSFWorkbook(outPath1);
        let* (_, ans)   = apply1(ma, wb as ##org.apache.poi.ss.usermodel.Workbook, None, Map.empty());
        let* _          = writeHSSFWorkbook(wb, outPath1);
        let* _          = closeHSSFWorkbook(wb);
        Ok(ans)

    def apply1(ma: SheetWriter[a], 
               wb: ##org.apache.poi.ss.usermodel.Workbook, 
               style: Option[CellStyle],
               s1: Cursors): Result[(Cursors, a), String] & Impure = 
        let SheetWriter(x1) = ma;
        x1(wb, style, s1)

    pub def askStyle(): SheetWriter[Option[CellStyle]] = 
        SheetWriter((_, style, s1) -> Ok((s1, style)) as & Impure)


    pub def point(x: a): SheetWriter[a] = 
        SheetWriter((_, _, s1) -> Ok((s1, x)) as & Impure)

    pub def throwError(msg: String): SheetWriter[a] =
        SheetWriter((_, _, _) -> Err(msg) as & Impure)


    pub def catchError(ma: SheetWriter[a], handle: String -> SheetWriter[a]): SheetWriter[a] =
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => Ok((s2, a))
            case Err(e) => apply1(handle(e), wb, style, s1)
        })

    pub def mapError(f: String -> String, ma: SheetWriter[a]): SheetWriter[a] =
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => Ok((s2, a))
            case Err(e) => Err(f(e))
        })

    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: SheetWriter[a]): SheetWriter[b] & ef = 
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => { let ans = fn(a) as & Impure; Ok((s2, ans)) }
            case Err(e) => Err(e)
        }) as & ef

    pub def ap(mf: SheetWriter[a -> b & ef], ma: SheetWriter[a]): SheetWriter[b] & ef = 
        SheetWriter((wb, style, s1) -> match apply1(mf, wb, style, s1) { 
            case Ok((s2, f)) => match apply1(ma, wb, style, s2) { 
                case Ok((s3, a)) => { let ans = f(a) as & Impure; Ok((s3, ans)) }
                case Err(e) => Err(e)
            }
            case Err(e) => Err(e)
        }) as & ef


    pub def bind(ma: SheetWriter[a], fn: a -> SheetWriter[b] & ef): SheetWriter[b] & ef = 
        SheetWriter((wb, style, s1) -> match apply1(ma, wb, style, s1) { 
            case Ok((s2, a)) => apply1(fn(a), wb, style, s2)
            case Err(e) => Err(e)
        }) as & ef

    pub def flatMap(fn: a -> SheetWriter[b] & ef, ma: SheetWriter[a]): SheetWriter[b] & ef = bind(ma, fn)


    pub def guard(x: Bool): SheetWriter[Unit] = 
        match x { 
            case true => point()
            case false => throwError("guard")
        }

    pub def whenTrue(x: Bool, ma: SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case true => ma
            case false => point()
        }

    pub def whenFalse(x: Bool, ma: SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case true => point()
            case false => ma
        }

    pub def whenSome(x: Option[a], f: a -> SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case Some(a) => f(a)
            case None => point()
        }

    pub def whenNone(x: Option[a], ma: SheetWriter[Unit]): SheetWriter[Unit] = 
        match x { 
            case Some(_) => point()
            case None => ma
        }



    def liftWorkbookPrim(fn: ##org.apache.poi.ss.usermodel.Workbook -> Result[a, String] & Impure): SheetWriter[a] = 
        SheetWriter((wb, _, s1) -> {let ans = fn(wb); Result.map(x -> (s1, x), ans)})

    pub def liftAction(fn: Unit -> a & Impure): SheetWriter[a] = 
        SheetWriter((_, _, s1)-> {let a = fn(); Ok((s1, a)) })

    pub def liftActionResult(fn: Unit -> Result[a, String] & Impure): SheetWriter[a] = 
        SheetWriter((_, _, s1) -> {let ans = fn(); Result.map(x -> (s1, x), ans)})

    pub def liftOption(x: Option[a], errMsg: String): SheetWriter[a] = match x { 
        case Some(x1) => point(x1)
        case None => throwError(errMsg)
    }

    pub def liftResult(x: Result[a, String]): SheetWriter[a] = match x { 
        case Ok(x1)   => point(x1)
        case Err(msg) => throwError(msg)
    }

    pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): SheetWriter[a] = match x { 
        case Ok(x1) => point(x1)
        case Err(e) => throwError(makeErrMsg(e))
    }

    pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetWriter[a] = match x { 
        case Success(x1) => point(x1)
        case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    }

    pub def localCellStyle(style: CellStyle, mx: SheetWriter[a]): SheetWriter[a] = 
        SheetWriter((wb, _, s1) -> apply1(mx, wb, Some(style), s1))


    pub def mapM(mf: a -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] =
        SheetWriter((wb, style, s1) ->
            mapMHelper(wb, style, s1, mf, xs, e -> Err(e), (s2, ys) -> Ok(s2, ys))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   style: Option[CellStyle],
                   s1: Cursors,
                   mf: a -> SheetWriter[b], 
                   xs: List[a], 
                   fk: String -> Result[(Cursors, List[b]), String], 
                   sk: (Cursors, List[b]) -> Result[(Cursors, List[b]), String]): Result[(Cursors, List[b]), String] & Impure =
        match xs {
            case Nil => sk(s1, Nil)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, x1)) => mapMHelper(wb, style, s2, mf, rs, fk, (s3, ks) -> sk(s3, x1 :: ks))
                }
            }
        }

    pub def mapX(mf: a -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] =
        SheetWriter((wb, style, s1) ->
            mapXHelper(wb, style, s1, mf, xs, e -> Err(e), s2 -> Ok((s2, ())))
        )

    def mapXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                   style: Option[CellStyle],
                   s1: Cursors,
                   mf: a -> SheetWriter[b], 
                   xs: List[a], 
                   fk: String -> Result[(Cursors, Unit), String], 
                   sk: Cursors -> Result[(Cursors, Unit), String]): Result[(Cursors, Unit), String] & Impure =
        match xs {
            case Nil => sk(s1)
            case x :: rs => {
                let ma = mf(x);
                match apply1(ma, wb, style, s1) {
                    case Err(e) => fk(e)
                    case Ok((s2, _)) => mapXHelper(wb, style, s2, mf, rs, fk, sk)
                }
            }
        }

    pub def mapWithIndexM(mf: (a, Int32) -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] =
        SheetWriter((wb, style, s1) ->
            mapWithIndexMHelper(wb, style, s1, mf, 0, xs, e -> Err(e), (s2, ys) -> Ok((s2, ys)))
        )

    ///
    /// Helper function for `mapM`.
    ///
    def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            style: Option[CellStyle],
                            s1: Cursors,
                            mf: (a, Int32) -> SheetWriter[b], 
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

    pub def mapWithIndexX(mf: (a, Int32) -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] =
        SheetWriter((wb, style, s1) ->
            mapWithIndexXHelper(wb, style, s1, mf, 0, xs, e -> Err(e), s2 -> Ok((s2, ())))
        )

    def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                            style: Option[CellStyle],
                            s1: Cursors,
                            mf: (a, Int32) -> SheetWriter[b],
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


    pub def sequence(xs: List[SheetWriter[a]]): SheetWriter[List[a]] =
        SheetWriter((wb, style, s1) ->
            sequenceHelper(wb, style, s1, xs, e -> Err(e), (s2, ks) -> Ok((s2, ks)))
        )

    ///
    /// Helper function for `sequence`.
    ///
    def sequenceHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                       style: Option[CellStyle],
                       s1: Cursors,
                       xs: List[SheetWriter[a]], 
                       fk: String -> Result[(Cursors, List[a]), String], 
                       sk: (Cursors, List[a]) -> Result[(Cursors, List[a]), String]): Result[(Cursors, List[a]), String] & Impure =
        match xs {
            case Nil => sk(s1, Nil)
            case mx :: rs => match apply1(mx, wb, style, s1) {
                case Err(e) => fk(e)
                case Ok((s2, x)) => sequenceHelper(wb, style, s2, rs, fk, (s3, ks) -> sk(s3, x :: ks))
            }
        }

    ///
    /// Returns `(v1 :: v2 :: ... v :: vn)` if each of `f(xs_i)` is successful.
    /// Otherwise fails with the first error encountered.
    ///
    /// `traverse` is a synonym for `mapM`.
    ///
    pub def traverse(f: a -> SheetWriter[b], xs: List[a]): SheetWriter[List[b]] = mapM(f, xs)

    ///
    /// `traverseX` is a synonym for `mapX`.
    ///
    pub def traverseX(f: a -> SheetWriter[b], xs: List[a]): SheetWriter[Unit] = mapX(f, xs)

    ///
    /// `forM` is `mapM` with its arguments flipped.
    ///
    pub def forM(xs: List[a], f: a -> SheetWriter[b]): SheetWriter[List[b]] = mapM(f, xs)

    ///
    /// `forX` is `mapX` with its arguments flipped.
    ///
    pub def forX(xs: List[a], f: a -> SheetWriter[b]): SheetWriter[Unit] = mapX(f, xs)


    pub def foldM(mf: (b, a) -> SheetWriter[b], start: b, xs: List[a]): SheetWriter[b] =
        SheetWriter((wb, style, s1) ->
            foldMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac) -> Ok((s2, ac)))
        )

    def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
                    style: Option[CellStyle],
                    s1: Cursors,
                    mf: (b, a) -> SheetWriter[b], 
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

    pub def mapAccumM(mf: (ac, a) -> SheetWriter[(ac, b)], start: ac, xs: List[a]): SheetWriter[(ac, List[b])] =
        SheetWriter((wb, style, s1) ->
            mapAccumMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac, ks) -> Ok((s2, (ac, ks))))
        )

    def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
                        style: Option[CellStyle],
                        s1: Cursors,
                        mf: (ac, a) -> SheetWriter[(ac, b)],
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
    pub def mapAccumX(mf: (ac, a) -> SheetWriter[ac], start: ac, xs: List[a]): SheetWriter[ac] = foldM(mf, start, xs)



    // SheetWriter

    def getCursorAndIncr(sheetName: String): SheetWriter[Int32] = 
        SheetWriter((_, _, s1) -> match Map.get(sheetName, s1) {
            case Some(ix)   => Ok((Map.adjust(x -> x + 1, sheetName, s1), ix)) as & Impure
            case None       => Ok((Map.insert(sheetName, 1, s1), 0)) as & Impure
        })

    

    pub def createSheet(name: String): SheetWriter[Sheet] =
        use SheetIO/Internal/Workbook.{createSheet};
        liftWorkbookPrim(x1 -> Result.map(Sheet, createSheet(x1, name))) 

    
    
    /// Although this is `SheetWriter` some introspection (reading) is needed, 
    /// e.g. if we are writing to a template file

    pub def getSheetNamed(name: String): SheetWriter[Sheet] =
        use SheetIO/Internal/Workbook.{getSheet};
        liftWorkbookPrim(x1 -> Result.map(Sheet, getSheet(x1, name))) 
    
    pub def getSheetNames(): SheetWriter[List[String]] =
        let* n  = getNumberOfSheets();
        let ixs = List.range(0, n);
        mapM(getSheetNameWb, ixs)      
    
    def getNumberOfSheets(): SheetWriter[Int32] =
        use SheetIO/Internal/Workbook.{getNumberOfSheets};
        liftWorkbookPrim(getNumberOfSheets)

    def getSheetNameWb(ix: Int32): SheetWriter[String] =
        use SheetIO/Internal/Workbook.{getSheetName};
        liftWorkbookPrim(x1 -> getSheetName(x1, ix))


    pub def createDataFormat(): SheetWriter[DataFormat] =
        use SheetIO/Internal/Workbook.{createDataFormat};
        liftWorkbookPrim(x1 -> Result.map(DataFormat, createDataFormat(x1)))

    pub def getFormat(x: DataFormat, format: String): SheetWriter[Int16] =
        use SheetIO/Internal/Workbook.{getFormat};
        let DataFormat(x1) = x;
        liftActionResult(_ -> getFormat(x1, format))


    pub def setCursor!(x: Sheet, n: Int32): SheetWriter[Unit] =
        let* name = getSheetName(x);
        setCursorHelper(name, n)

    def getSheetName(x: Sheet): SheetWriter[String] =
        use SheetIO/Internal/Sheet.{getSheetName};
        let Sheet(x1) = x;
        liftActionResult(_ -> getSheetName(x1))

    def setCursorHelper(sheetName: String, n: Int32): SheetWriter[Unit] = 
        SheetWriter((_, _, s1) -> Ok((Map.insert(sheetName, n, s1), ())) as & Impure)


    pub def createCellStyle(mx: CellStyleSetter[a]): SheetWriter[CellStyle] =
        use SheetIO/Internal/Workbook.{createCellStyle};
        let CellStyleSetter(mf) = mx;
        let* style              = liftWorkbookPrim(createCellStyle >> Result.map(CellStyle));
        let CellStyle(style1)   = style;
        let* _                  = liftActionResult(_ -> mf(style1));
        point(style)



    pub def createFont(mx: FontSetter[a]): SheetWriter[Font] =
        use SheetIO/Internal/Workbook.{createFont};
        let FontSetter(mf)      = mx;
        let* font               = liftWorkbookPrim(createFont >> Result.map(Font));
        let Font(font1)         = font;
        let* _                  = liftActionResult(_ -> mf(font1));
        point(font)

    // createRow
    def createRow(x: Sheet, ix: Int32): SheetWriter[Row] =
        use SheetIO/Internal/Sheet.{createRow};
        let Sheet(x1) = x;
        liftActionResult(_ -> Result.map(Row, createRow(x1, ix)))

    pub def appendRow(x: Sheet, mx: RowSetter[a]): SheetWriter[Unit] =
        let RowSetter(mf) = mx;
        let* name       = getSheetName(x);
        let* ix         = getCursorAndIncr(name);
        let* row        = createRow(x, ix);
        let Row(row1)   = row;
        let* style      = askStyle();
        let* _          = liftActionResult(_ -> mf(row1, style));
        point()


    pub def allRows(sheet: Sheet, rows: List[a], writer: a -> RowSetter[Unit]): SheetWriter[Unit] = 
        mapX(r -> appendRow(sheet, writer(r)), rows)


    /// Row 0 is read for the headers, all subsequent rows are *data*.
    pub def allRowsWithHeaders(sheet: Sheet, headers: List[String], rows: List[a], writer: a -> RowSetter[Unit]): SheetWriter[Unit] = 
        let* _      = appendRow(sheet, SheetIO/RowSetter.putHeaders!(headers));
        let* _      = mapX(r -> appendRow(sheet, writer(r)), rows);
        point()

}
