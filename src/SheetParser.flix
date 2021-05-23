/*
 * Copyright 2021 Stephen Tetley
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



namespace SheetIO/SheetParser {

    type alias Sheet = ##org.apache.poi.ss.usermodel.Sheet

    type alias SourcePosition = {row: Int32, column: Int32}

    type alias ParseState[st] = {row: Int32, column: Int32, userState: st}

    type alias Failures = Chain[String]

    type alias FailK[ka1] = SourcePosition -> Failures ~> ka1

    type alias SuccessK[ka1, st, ans] = ans -> FailK[ka1] -> Sheet -> ParseState[st] -> Failures ~> ka1


    /// `st` is the type of user state
    /// `ans` is the type of "immediate answer"
    /// `ka1` is the continuation answer type - if Flix had user visible existentials it would 
    /// be hidden as an existential type. Generally it will always be represented with a type variable
    /// except in "run" functions.
    pub opaque type GenSheetParser[ka1, st, ans] = SuccessK[ka1, st, ans] -> FailK[ka1] -> Sheet -> ParseState[st] -> Failures ~> ka1

    type alias SheetParser[ka1, ans] = GenSheetParser[ka1, Unit, ans]

    type alias InternalError = {position: SourcePosition, messages: Chain[String]}
    
    ///
    /// This is exposed so client code can implement wrappers over `run` etc. without
    /// knowing about the internal error type `InternalError`.
    ///
    type alias GenSheetParser1[st, ans] = GenSheetParser[Result[ans, InternalError], st, ans]

    type alias SheetParser1[ans] = SheetParser[Result[ans, InternalError], ans]

    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXlsx(p: GenSheetParser1[st, a], sheetName: String, state: st, path: System.FilePath): Result[a, String] & Impure = 
        runXssf(p, sheetName, state, path)
    
    /// Excel 2007 OOXML (.xlsx) format.
    pub def runXssf(p: GenSheetParser1[st, a], sheetName: String, state: st, path: System.FilePath): Result[a, String] & Impure = 
        use Result.{flatMap};
        let path1 = ToString.toString(path);
        match SheetIO/Internal/Prim.openXSSFWorkbook(path1) {
            case Err(e) => Err("runXssf('${path1}'): ${e}")
            case Ok(wb) => {
                let* sheet = SheetIO/Internal/Workbook.getSheet(wb as ##org.apache.poi.ss.usermodel.Workbook, sheetName);
                let state = {row = 0, column = 0, userState = state};
                let fk = (posn, errs) -> Err({messages = errs, position = posn}) as & Impure;
                let sk = (a, _, _, _, _) -> Ok(a) as & Impure;
                let GenSheetParser(p1) = p;
                let ans = p1(sk, fk, sheet, state, Chain.empty());
                let* _ = SheetIO/Internal/Prim.closeXSSFWorkbook(wb);
                match ans {
                    case Ok(a) => Ok(a)
                    case Err(e1) => Err("Error {row: ${e1.position.row}, column: ${e1.position.column}}")
                }
            }
        }


    pub def return(x: a): GenSheetParser[ka1, st, a] = GenSheetParser(sk -> sk(x))    

    pub def throwError(message: String) : GenSheetParser[ka1, st, ans] =
        use Chain.{singleton};
        GenSheetParser((_, fk, _, st, _) -> fk({row = st.row, column = st.column}, singleton(message)))

    pub def fail(message: String) : GenSheetParser[ka1, st, ans] = throwError(message)

    def apply1 (p: GenSheetParser[ka1, st, ans], 
                sk: SuccessK[ka1, st, ans], 
                fk: FailK[ka1], 
                sheet: Sheet, 
                st: ParseState[st], 
                errs: Failures): ka1 & Impure =
        let GenSheetParser(pf) = p;
        pf(sk, fk, sheet, st, errs)


    ///
    /// `swapError` is quite subtle. Swapping the error continuation needs to be
    /// a "local" operation which is undone (restored) afterwards (c.f. local in the Reader
    /// monad). It must not be a "stateful" operation with unlimited extent.
    ///
    pub def swapError(p: GenSheetParser[ka1, st, ans], message: String) : GenSheetParser[ka1, st, ans] =
        use Chain.{singleton};
        GenSheetParser((sk, fk, sheet, st, errs) ->
            let fk1 = (pos, _) -> fk(pos, singleton(message));
            let sk1 = (x, _, _, st1, _) -> sk(x, fk, sheet, st1, errs);
            apply1(p, sk1, fk1, sheet, st, errs)
        )


    pub def map(f: a -> b & ef, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        GenSheetParser((sk, fk, sheet, pos, errs) ->
            let sk1 = (x, fk1, _, pos1, errs1) -> sk(f(x) as & Impure, fk1, sheet, pos1, errs1);
            apply1(p, sk1, fk, sheet, pos, errs)
        )

    pub def ap(mf: GenSheetParser[ka1, st, a -> b & ef], ma: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        let* f = mf;
        let* a = ma;
        liftAction(_ -> f(a))

    pub def bind(ma: GenSheetParser[ka1, st, a], k: a -> GenSheetParser[ka1, st, b]):  GenSheetParser[ka1, st, b] =
        GenSheetParser(sk ->
            apply1(ma, x -> apply1(k(x), sk))
        )

    pub def flatMap(k: a -> GenSheetParser[ka1, st, b], ma: GenSheetParser[ka1, st, a]):  GenSheetParser[ka1, st, b] = bind(ma, k)


    pub def >>=(ma: GenSheetParser[ka1, st, a], k: a -> GenSheetParser[ka1, st, b]):  GenSheetParser[ka1, st, b] =
        bind(ma, k)

    pub def liftAction(f: Unit -> a & ef): GenSheetParser[ka1, st, a] =
        GenSheetParser((sk, fk, sheet, st, errs) ->
            let a = f() as & Impure;
            sk(a, fk, sheet, st, errs)
        )
        
    pub def liftSheetPrim(f: Sheet ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        use Chain.{singleton};
        GenSheetParser((sk, fk, sheet, st, errs) ->
            let ans = f(sheet);
            match ans { 
                case Ok(a) => sk(a, fk, sheet, st, errs)
                case Err(msg) => fk({row = st.row, column = st.column}, singleton("liftSheetPrim - ${msg}"))
            }
        )

    pub def liftRowPrim(f: ##org.apache.poi.ss.usermodel.Row ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        getRowIndex() >>= ix -> liftRowPrimHelper(ix, f)

    def liftRowPrimHelper(rowIx: Int32, 
                        f: ##org.apache.poi.ss.usermodel.Row ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        use Result.{flatMap};
        liftSheetPrim(sheet -> {
            let* rmax = SheetIO/Internal/Sheet.getLastRowNum(sheet);
            if (rowIx > rmax) 
                Err("At end of Sheet")
            else {
                let* row = SheetIO/Internal/Sheet.getRow(sheet, rowIx);
                f(row)
            }
        })

    pub def liftCellPrim(f: ##org.apache.poi.ss.usermodel.Cell ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        getColumnIndex() >>= ix -> liftCellPrimHelper(ix, f)

    def liftCellPrimHelper(cellIx: Int32, 
                            f: ##org.apache.poi.ss.usermodel.Cell ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        use Result.{flatMap};
        liftRowPrim(row -> {
            let* cmax = SheetIO/Internal/Row.getLastCellNum(row);
            if (cellIx > cmax) 
                Err("At end of Row")
            else {
                let* cell = SheetIO/Internal/Row.getCell(row, cellIx);
                f(cell)
            }
        })

    pub def ifThenElse(cond: GenSheetParser[ka1, st, a], 
                        thenp: a -> GenSheetParser[ka1, st, b], 
                        elsep: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, b] = 
        GenSheetParser((sk, fk, sheet, st) ->
            apply1(cond, x -> apply1(thenp(x), sk), _ -> apply1(elsep, sk, fk, sheet, st), sheet, st))

    
    def getRowIndex(): GenSheetParser[ka1, st, Int32] =  
        GenSheetParser((sk, fk, sheet, st) ->
            sk(st.row, fk, sheet, st)
        )
    
    def getColumnIndex(): GenSheetParser[ka1, st, Int32] =  
        GenSheetParser((sk, fk, sheet, st) ->
            sk(st.column, fk, sheet, st)
        )

    pub def getSourcePosition(): GenSheetParser[ka1, st, SourcePosition] =  
        GenSheetParser((sk, fk, sheet, st) ->
            sk({row = st.row, column = st.column}, fk, sheet, st)
        )

    pub def setSourcePos(pos: SourcePosition): GenSheetParser[ka1, st, Unit] =  
        GenSheetParser((sk, fk, sheet, st) ->
            let st1 = {row = pos.row, column = pos.column | st};
            sk((), fk, sheet, st1)
        ) 

    pub def getState(): GenSheetParser[ka1, st, st] =  
        GenSheetParser((sk, fk, sheet, st) ->            
            sk(st.userState, fk, sheet, st)
        )

    pub def setState(state: st): GenSheetParser[ka1, st, Unit] =  
        GenSheetParser((sk, fk, sheet, st) ->
            let st1 = {userState = state | st};
            sk((), fk, sheet, st1)
        )


    pub def updateState(modify: st -> st): GenSheetParser[ka1, st, Unit] =  
        GenSheetParser((sk, fk, sheet, st) ->
            let st1 = {userState = modify(st.userState) | st};
            sk((), fk, sheet, st1)
        )  

    /// TODO - maybe incrementing row should be manual...
    def incSourcePos(): GenSheetParser[ka1, st, Unit] =  
        let* pos = getSourcePosition();
        /// cmax is +1 the zero-indexed column number...
        let* cmax = liftRowPrim(SheetIO/Internal/Row.getLastCellNum);
        if (pos.column >= cmax)
            setSourcePos({row = pos.row + 1, column = 0})
        else 
            setSourcePos({column = pos.column + 1 | pos})


    pub def getContent(): GenSheetParser[ka1, st, String] =  
        let* ans    = liftCellPrim(SheetIO/Internal/Cell.getDefaultFormatterCellValue);
        let* _      = incSourcePos();
        return(ans)

}

