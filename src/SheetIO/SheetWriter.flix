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


    /// Imperative writing is not a writer monad (which has backtracking)
    /// Imperative writing uses mutable state

    pub type alias SheetWriter[ka: Type, r: Region, ef: Eff, a: Type] =
        Graded.EnvMutStateError[ka, Option[Internal.Apache.Poi.SS.Usermodel.CellStyle[r]], Internal.Apache.Poi.SS.Usermodel.MutSheet[r], ef, a]
    

}

mod SheetIO.SheetWriter {

    use GradedApplicative.{point}
    use GradedMonad.{>>=}
    use Graded.EnvMutStateError.{liftSetter!, liftSetterResult!, liftGetter, liftGetterResult}

    use SheetIO.SheetWriter
    use SheetIO.RowWriter
    
    use Internal.Apache.Poi.SS.Usermodel.CellStyle
    use Internal.Apache.Poi.SS.Usermodel.MutRow
    use Internal.Apache.Poi.SS.Usermodel.MutSheet



    /// The state param `st` is expected to be mutated...
    pub def runSheetWriter(ma: SheetWriter[a, r, ef, a],
                            env: Option[CellStyle[r]],
                            st: MutSheet[r]): Result[String, a] \ ef =
        Graded.EnvMutStateError.runEnvMutStateError(ma, env, st)

    pub def withRowWriter(write: RowWriter[a, r, ef, a], row: MutRow[r]): SheetWriter[ka, r, ef, a] =
        liftSetterResult!((_, _) -> SheetIO.RowWriter.runRowWriter(write, row))



    def getRow(ix: Int32): SheetWriter[ka, r, r, Option[MutRow[r]]] = 
        liftGetterResult((_, sheet) -> Internal.Apache.Poi.SS.Usermodel.MutSheet.getRow(ix, sheet)) 
            |> GradedError.reifyError
            |> GradedFunctor.map(Result.toOption)

    def createRow!(ix: Int32): SheetWriter[ka, r, r, MutRow[r]] = 
        liftSetter!((_, sheet) -> Internal.Apache.Poi.SS.Usermodel.MutSheet.createRow!(ix, sheet))

    def getOrCreateRow!(ix: Int32): SheetWriter[ka, r, r, MutRow[r]] = 
        getRow(ix) >>= ans -> 
        match ans {
            case Some(row)  => point(row)
            case None       => createRow!(ix)
        }

    // call `getOrCreateRow!` first ...
    pub def writeAt!(ix: Int32, write: RowWriter[a, r, ef, a]): SheetWriter[ka, r, ef + r, a] =
        getOrCreateRow!(ix)    >>= row -> 
        withRowWriter(write, row)


    pub def getSheetName(): SheetWriter[ka, r, r, String] =
        liftSetter!((_, sheet) -> Internal.Apache.Poi.SS.Usermodel.MutSheet.getSheetName(sheet))



/*


    // pub def write(ma: SheetWriter[r, ka, a], sheet: Sheet[r]): Result[a, EvalError] & r = 
    //     let Sheet(sheet1) = sheet;
    //     apply1(ma, None, sheet1)


    pub def askStyle(): SheetWriter[r, ka, Option[CellStyle[r]]] = 
        askEnv()


    // pub def liftResult1(x: Result[a, e], makeErrMsg: e -> String): SheetWriter[r, ka, a] = match x { 
    //     case Ok(x1) => point(x1)
    //     case Err(e) => throwError(makeErrMsg(e))
    // }

    // pub def liftValidation(x: Validation[a, e], makeErrMsg: List[e] -> String): SheetWriter[r, ka, a] = match x { 
    //     case Success(x1) => point(x1)
    //     case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    // }

    pub def localCellStyle(style: CellStyle, mx: SheetWriter[r, ka, a]): SheetWriter[r, ka, a] = 
        SheetWriter((wb, _, s1) -> apply1(mx, wb, Some(style), s1))



    // SheetWriter

    def getCursorAndIncr(sheetName: String): SheetWriter[r, ka, Int32] = 
        SheetWriter((_, _, s1) -> match Map.get(sheetName, s1) {
            case Some(ix)   => Ok((Map.adjust(x -> x + 1, sheetName, s1), ix)) as \ IO
            case None       => Ok((Map.insert(sheetName, 1, s1), 0)) as \ IO
        })

    


    
    
    /// Although this is `SheetWriter` some introspection (reading) is needed, 
    /// e.g. if we are writing to a template file





    pub def getFormat(x: DataFormat, format: String): SheetWriter[r, ka, Int16] =
        use SheetIO.Internal.Workbook.{getFormat};
        let DataFormat(x1) = x;
        liftActionResult(_ -> getFormat(x1, format))


    def setCursorHelper(sheetName: String, n: Int32): SheetWriter[r, ka, Unit] = 
        SheetWriter((_, _, s1) -> Ok((Map.insert(sheetName, n, s1), ())) as \ IO)



    // createRow
    def createRow!(ix: Int32): SheetWriter[r, ka, Row[r]] \ r =
        liftSetter!((_, sheet) -> Internal.Apache.Poi.SS.Usermodel.Sheet.createRow!(ix, sheet))


    pub def setRow!(ix: Int32, ma: RowWriter[r, a, a]): SheetWriter[r, ka, a] \ r =
        let* row        = createRow!(ix);
        let r           = Scoped.regionOf(row);
        let* style      = askStyle();
        liftActionResult(_ -> SheetIO.RowWriter.runRowWriter!(r, ma, style, row)) 



    pub def allRows!(writer: a -> RowWriter[r, Unit], rows: List[a]): SheetWriter[r, ka, Unit] \ r = 
        FoldableWithIndex.foreachMWithIndex((ix, a) -> setRow!(ix, writer(a)), rows)

    pub def allRowsForm!(start: Int32, writer: a -> RowWriter[r, b], rows: List[a]): SheetWriter[r, ka, Unit] \ r = 
        FoldableWithIndex.foreachMWithIndex((ix, a) -> setRow!(ix + start, writer(a)), rows)


    /// Headers are written to Row 0, all subsequent rows are *data*.
    pub def allRowsWithHeaders!(headers: List[String], writer: a -> RowWriter[r, b], rows: List[a]): SheetWriter[r, ka, Unit] \ r  = 
        let* _      = setRow!(0, SheetIO.RowWriter.putHeaders!(headers));
        let* _      = allRowsForm!(1, writer, rows);
        point()


*/

/*
    // pub def mapWithIndexM(mf: (a, Int32) -> SheetWriter[r, ka, b], xs: List[a]): SheetWriter[r, ka, List[b]] =
    //     SheetWriter((wb, style, s1) ->
    //         mapWithIndexMHelper(wb, style, s1, mf, 0, xs, e -> Err(e), (s2, ys) -> Ok((s2, ys)))
    //     )

    // ///
    // /// Helper function for `mapM`.
    // ///
    // def mapWithIndexMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
    //                         style: Option[CellStyle],
    //                         s1: Cursors,
    //                         mf: (a, Int32) -> SheetWriter[r, ka, b], 
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

    // pub def mapWithIndexX(mf: (a, Int32) -> SheetWriter[r, ka, b], xs: List[a]): SheetWriter[r, ka, Unit] =
    //     SheetWriter((wb, style, s1) ->
    //         mapWithIndexXHelper(wb, style, s1, mf, 0, xs, e -> Err(e), s2 -> Ok((s2, ())))
    //     )

    // def mapWithIndexXHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
    //                         style: Option[CellStyle],
    //                         s1: Cursors,
    //                         mf: (a, Int32) -> SheetWriter[r, ka, b],
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



    // ///
    // /// `forM` is `mapM` with its arguments flipped.
    // ///
    // pub def forM(xs: List[a], f: a -> SheetWriter[r, ka, b]): SheetWriter[r, ka, List[b]] = mapM(f, xs)

    // ///
    // /// `forX` is `mapX` with its arguments flipped.
    // ///
    // pub def forX(xs: List[a], f: a -> SheetWriter[r, ka, b]): SheetWriter[r, ka, Unit] = mapX(f, xs)


    // pub def foldM(mf: (b, a) -> SheetWriter[r, ka, b], start: b, xs: List[a]): SheetWriter[r, ka, b] =
    //     SheetWriter((wb, style, s1) ->
    //         foldMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac) -> Ok((s2, ac)))
    //     )

    // def foldMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook, 
    //                 style: Option[CellStyle],
    //                 s1: Cursors,
    //                 mf: (b, a) -> SheetWriter[r, ka, b], 
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

    // pub def mapAccumM(mf: (ac, a) -> SheetWriter[r, ka, (ac, b)], start: ac, xs: List[a]): SheetWriter[r, ka, (ac, List[b])] =
    //     SheetWriter((wb, style, s1) ->
    //         mapAccumMHelper(wb, style, s1, mf, start, xs, e -> Err(e), (s2, ac, ks) -> Ok((s2, (ac, ks))))
    //     )

    // def mapAccumMHelper(wb: ##org.apache.poi.ss.usermodel.Workbook,
    //                     style: Option[CellStyle],
    //                     s1: Cursors,
    //                     mf: (ac, a) -> SheetWriter[r, ka, (ac, b)],
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
    // pub def mapAccumX(mf: (ac, a) -> SheetWriter[r, ka, ac], start: ac, xs: List[a]): SheetWriter[r, ka, ac] = foldM(mf, start, xs)


*/

}
