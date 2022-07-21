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
    use SheetIO/Internal/EvalError.{EvalError};
   

    pub enum SheetWriter[r: Region, a: Type](
        Option[CellStyle[r]] -> ##org.apache.poi.ss.usermodel.Sheet -> Result[a, EvalError] & r)


    instance Functor[SheetWriter[r]] {
        pub def map(f: a -> b & ef, o: SheetWriter[r, a]): SheetWriter[r, b] & ef = 
            SheetIO/SheetWriter.map(f, o)
    }

    instance Applicative[SheetWriter[r]] {
        pub def point(x: a): SheetWriter[r, a] = 
            SheetIO/SheetWriter.point(x)

        pub def ap(f: SheetWriter[r, a -> b & ef], x: SheetWriter[r, a]) : SheetWriter[r, b] & ef = 
            SheetIO/SheetWriter.ap(f, x)
    }

    instance Monad[SheetWriter[r]] {
        pub def flatMap(f: a -> SheetWriter[r, b] & ef, o: SheetWriter[r, a]): SheetWriter[r, b] & ef =
            SheetIO/SheetWriter.flatMap(f, o)
    }


}

namespace SheetIO/SheetWriter {

    use SheetIO.SheetWriter;
    use SheetIO.SheetWriter.{SheetWriter};
    
    use SheetIO.CellStyle;
    use SheetIO.CellStyle.{CellStyle};
    use SheetIO.Cursor;
    use SheetIO/Internal/EvalError.{EvalError};
    use SheetIO/Internal/PrimError.{PrimError};
    use SheetIO.Row;
    use SheetIO.Row.{Row};

    def apply1(ma: SheetWriter[r, a], 
               style: Option[CellStyle[r]], 
               sheet: ##org.apache.poi.ss.usermodel.Sheet): Result[a, EvalError] & r = 
        let SheetWriter(f1) = ma;
        f1(style, sheet)

    pub def askStyle(): SheetWriter[r, Option[CellStyle[r]]] = 
        SheetWriter((style, _) -> Ok(style) as & r) 



    pub def point(x: a): SheetWriter[r, a] = 
        SheetWriter((_, _) -> Ok(x) as & r)
/*
    pub def throwError(msg: String): SheetWriter[r, a] =
        SheetWriter((_, _, _) -> Err(msg) as & Impure)


    pub def catchError(ma: SheetWriter[r, a], handle: String -> SheetWriter[r, a]): SheetWriter[r, a] =
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
    pub def map(fn: a -> b & ef, ma: SheetWriter[r, a]): SheetWriter[r, b] & ef = 
        SheetWriter((style, sheet) -> 
            match apply1(ma, style, sheet) { 
                case Ok(a) => {fn(a) as & r} |> Ok
                case Err(e) => Err(e)
        }) as & ef


    pub def ap(mf: SheetWriter[r, a -> b & ef], ma: SheetWriter[r, a]): SheetWriter[r, b] & ef = 
        SheetWriter((style, sheet) -> 
            match apply1(mf, style, sheet) { 
                case Ok(f)  => match apply1(ma, style, sheet) { 
                    case Ok(a)  => {f(a) as & r} |> Ok
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        ) as & ef


    pub def flatMap(fn: a -> SheetWriter[r, b] & ef, ma: SheetWriter[r, a]): SheetWriter[r, b] & ef =
        SheetWriter((style, sheet) -> 
            match apply1(ma, style, sheet) { 
                case Ok(a)  => apply1(fn(a) as & r, style, sheet)
                case Err(e) => Err(e)
            }
        ) as & ef

    def liftSheetPrim(fn: Region[r] -> ##org.apache.poi.ss.usermodel.Sheet -> Result[a, PrimError] & ef): SheetWriter[r, a] & ef = 
        SheetWriter((_, sheet) -> 
            let r = () as Region[r];
            {fn(r, sheet) as & r} |> Result.mapErr(SheetIO/Internal/EvalError.emptyEvalError)
        )  as & ef

/*
    pub def guard(x: Bool): SheetWriter[r, Unit] = 
        match x { 
            case true => point()
            case false => throwError("guard")
        }

    pub def whenTrue(x: Bool, ma: SheetWriter[r, Unit]): SheetWriter[r, Unit] = 
        match x { 
            case true => ma
            case false => point()
        }

    pub def whenFalse(x: Bool, ma: SheetWriter[r, Unit]): SheetWriter[r, Unit] = 
        match x { 
            case true => point()
            case false => ma
        }

    pub def whenSome(x: Option[a], f: a -> SheetWriter[r, Unit]): SheetWriter[r, Unit] = 
        match x { 
            case Some(a) => f(a)
            case None => point()
        }

    pub def whenNone(x: Option[a], ma: SheetWriter[r, Unit]): SheetWriter[r, Unit] = 
        match x { 
            case Some(_) => point()
            case None => ma
        }



    def liftWorkbookPrim(fn: ##org.apache.poi.ss.usermodel.Workbook -> Result[a, String] & Impure): SheetWriter[r, a] = 
        SheetWriter((wb, _, s1) -> {let ans = fn(wb); Result.map(x -> (s1, x), ans)})

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




    pub def createDataFormat(): SheetWriter[r, DataFormat] =
        use SheetIO/Internal/Workbook.{createDataFormat};
        liftWorkbookPrim(x1 -> Result.map(DataFormat, createDataFormat(x1)))

    pub def getFormat(x: DataFormat, format: String): SheetWriter[r, Int16] =
        use SheetIO/Internal/Workbook.{getFormat};
        let DataFormat(x1) = x;
        liftActionResult(_ -> getFormat(x1, format))


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
*/
    // createRow
    def createRow(ix: Int32): SheetWriter[r, Row[r]] \ Write(r) =
        liftSheetPrim((r, sheet) -> SheetIO/Internal/Sheet.createRow(r, ix, sheet) |> Result.map(Row))
/*

    pub def appendRow(ix: Int32, mx: RowSetter[a]): SheetWriter[r, Unit] =
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
