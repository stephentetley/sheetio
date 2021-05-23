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

    /// `alt` rolls back the parser state (including the position) if `p` fails.
    pub def alt(p: GenSheetParser[ka1, st, a], q: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =
        ifThenElse(p, return, q)

    pub def seqLeft(p: GenSheetParser[ka1, st, a], q: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, a] =
        let* a = p;
        let* _ = q;
        return(a)

    pub def seqRight(p: GenSheetParser[ka1, st, a], q: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, b] =
        let* _ = p;
        let* b = q;
        return(b)

    pub def kleisliLeft(f1: a -> GenSheetParser[ka1, st, b], f2: b -> GenSheetParser[ka1, st, c], x: a): GenSheetParser[ka1, st, c] =
        let* b  = f1(x);
        let* c  = f2(b);
        return(c)

    pub def kleisliRight(f1: b -> GenSheetParser[ka1, st, c], f2: a -> GenSheetParser[ka1, st, b], x: a): GenSheetParser[ka1, st, c] =
        kleisliLeft(f2, f1, x)


    /// Alias for `seqRight`
    pub def *>(p1: GenSheetParser[ka1, st, a], p2: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, b] = 
        seqRight(p1, p2)

    /// Alias for `seqLeft`
    pub def <*(p1: GenSheetParser[ka1, st, a], p2: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, a] = 
        seqLeft(p1, p2)

    pub def <!>(p1: GenSheetParser[ka1, st, a], p2: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =
        alt(p1, p2)

    // Haskell's <$>
    pub def <&>(f: a -> b, p2: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        map(f, p2)

    pub def <&(x: b, p1: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        p1 *> return(x)

    pub def &>(p1: GenSheetParser[ka1, st, a], x: b): GenSheetParser[ka1, st, b] =
        p1 *> return(x)

    /// Flipped <&>
    pub def <&&>(p2: GenSheetParser[ka1, st, a], f: a -> b): GenSheetParser[ka1, st, b] =
        map(f, p2)


    // Haskell's ap <*>
    pub def <*>(mf: GenSheetParser[ka1, st, a -> b], p2: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        ap(mf, p2)

    /// Flipped <*>
    pub def <**>(p2: GenSheetParser[ka1, st, a], mf: GenSheetParser[ka1, st, a -> b]): GenSheetParser[ka1, st, b] =
        ap(mf, p2)        


    pub def >>=(p: GenSheetParser[ka1, st, a], k: a -> GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, b] =
        bind(p, k)


    pub def =<<(k: a -> GenSheetParser[ka1, st, b], p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        bind(p, k)

    pub def >=>(f1: a -> GenSheetParser[ka1, st, b], f2: b -> GenSheetParser[ka1, st, c], x: a): GenSheetParser[ka1, st, c] =
        kleisliLeft(f1, f2, x)

    pub def <=<(f1: b -> GenSheetParser[ka1, st, c], f2: a -> GenSheetParser[ka1, st, b], x: a): GenSheetParser[ka1, st, c] = 
        kleisliRight(f1, f2, x)

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
        GenSheetParser((sk, fk, sheet, st, errs) ->
            let idx = st.row;
            sk(idx, fk, sheet, st, errs)
        )
    
    def getColumnIndex(): GenSheetParser[ka1, st, Int32] =  
        GenSheetParser((sk, fk, sheet, st, errs) ->
            let idx = st.column;
            sk(idx, fk, sheet, st, errs)
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


    pub def tuple2(p1: GenSheetParser[ka1, st, a], p2: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, (a, b)] =
        let* a = p1; 
        let* b = p2;
        return((a, b))

    pub def tuple3(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c]): GenSheetParser[ka1, st, (a, b, c)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        return((a, b, c))

    pub def tuple4(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d]): GenSheetParser[ka1, st, (a, b, c, d)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        return((a, b, c, d))


    pub def tuple5(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d], 
                    p5: GenSheetParser[ka1, st, e]): GenSheetParser[ka1, st, (a, b, c, d, e)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        let* e = p5;
        return((a, b, c, d, e))

    pub def tuple6(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d], 
                    p5: GenSheetParser[ka1, st, e], 
                    p6: GenSheetParser[ka1, st, f]): 
                    GenSheetParser[ka1, st, (a, b, c, d, e, f)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        let* e = p5;
        let* f = p6;
        return((a, b, c, d, e, f))

    pub def tuple7(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d], 
                    p5: GenSheetParser[ka1, st, e], 
                    p6: GenSheetParser[ka1, st, f], 
                    p7: GenSheetParser[ka1, st, g]): GenSheetParser[ka1, st, (a, b, c, d, e, f, g)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        let* e = p5;
        let* f = p6;
        let* g = p7;
        return((a, b, c, d, e, f, g))

    pub def tuple8(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d], 
                    p5: GenSheetParser[ka1, st, e], 
                    p6: GenSheetParser[ka1, st, f], 
                    p7: GenSheetParser[ka1, st, g], 
                    p8: GenSheetParser[ka1, st, h]): GenSheetParser[ka1, st, (a, b, c, d, e, f, g, h)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        let* e = p5;
        let* f = p6;
        let* g = p7;
        let* h = p8;
        return((a, b, c, d, e, f, g, h))

    pub def tuple9(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d], 
                    p5: GenSheetParser[ka1, st, e], 
                    p6: GenSheetParser[ka1, st, f], 
                    p7: GenSheetParser[ka1, st, g], 
                    p8: GenSheetParser[ka1, st, h], 
                    p9: GenSheetParser[ka1, st, i]): GenSheetParser[ka1, st, (a, b, c, d, e, f, g, h, i)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        let* e = p5;
        let* f = p6;
        let* g = p7;
        let* h = p8;
        let* i = p9;
        return((a, b, c, d, e, f, g, h, i))

    pub def tuple10(p1: GenSheetParser[ka1, st, a], 
                    p2: GenSheetParser[ka1, st, b], 
                    p3: GenSheetParser[ka1, st, c], 
                    p4: GenSheetParser[ka1, st, d], 
                    p5: GenSheetParser[ka1, st, e], 
                    p6: GenSheetParser[ka1, st, f], 
                    p7: GenSheetParser[ka1, st, g], 
                    p8: GenSheetParser[ka1, st, h], 
                    p9: GenSheetParser[ka1, st, i], 
                    p10: GenSheetParser[ka1, st, j]): GenSheetParser[ka1, st, (a, b, c, d, e, f, g, h, i, j)] =
        let* a = p1; 
        let* b = p2;
        let* c = p3;
        let* d = p4;
        let* e = p5;
        let* f = p6;
        let* g = p7;
        let* h = p8;
        let* i = p9;
        let* j = p10;
        return((a, b, c, d, e, f, g, h, i, j))

    /// TODO - maybe incrementing row should be manual...
    def incColumnPos(): GenSheetParser[ka1, st, Unit] =  
        let* pos = getSourcePosition();
        setSourcePos({column = pos.column + 1 | pos})
        
    def readCellWith(fn: ##org.apache.poi.ss.usermodel.Cell ~> Result[a, String]): GenSheetParser[ka1, st, a] = 
        let* ans    = liftCellPrim(fn);
        let* _      = incColumnPos();
        return(ans)

    pub def content(): GenSheetParser[ka1, st, String] =  
        readCellWith(SheetIO/Internal/Cell.getDefaultFormatterCellValue)
        
    pub def boolean(): GenSheetParser[ka1, st, Bool] =
        readCellWith(SheetIO/Internal/Cell.getBooleanCellValue)

    pub def numeric(): GenSheetParser[ka1, st, Float64] =
        readCellWith(SheetIO/Internal/Cell.getNumericCellValue)

    pub def literal(s: String): GenSheetParser[ka1, st, String] =
        let* s1 = content();
        if (s == s1) 
            return(s)
        else 
            throwError("literal - no match")

    pub def row(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =  
        let* ans    = p;
        let* pos    = getSourcePosition();
        let* _      = setSourcePos({row = pos.row + 1, column = 0});
        return(ans)

    ///
    /// lookahead - run the parser but don't consume input on success
    /// On fail, fail.
    ///
    pub def lookahead(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =
        GenSheetParser((sk, fk, sheet, st) ->
            let sk1 = (ans, fk1, _, _) -> sk(ans, fk1, sheet, st);
            apply1(p, sk1, fk, sheet, st)
        )

    pub def endOfInput() : GenSheetParser[ka1, st, Unit] =
        GenSheetParser((sk, fk, sheet, st, errs) -> 
            match SheetIO/Internal/Sheet.getLastRowNum(sheet) {
                case Err(msg) => fk({row = st.row, column = st.column}, Chain.cons("endOfInput - ${msg}", errs))
                case Ok(rmax) => 
                    if (st.row >= rmax) 
                        sk((), fk, sheet, st, errs)
                    else 
                        fk({row = st.row, column = st.column}, Chain.cons("endOfInput - not at end", errs))
            }
        )

    pub def isAtEnd() : GenSheetParser[ka1, st, Bool] =
        GenSheetParser((sk, fk, sheet, st, errs) -> 
            match SheetIO/Internal/Sheet.getLastRowNum(sheet) {
                case Err(msg) => fk({row = st.row, column = st.column}, Chain.cons("isEof - ${msg}", errs))
                case Ok(rmax) => 
                    if (st.row >= rmax) 
                        sk(true, fk, sheet, st, errs)
                    else 
                        sk(false, fk, sheet, st, errs)
            }
        )

    pub def satisfy(f: String -> Bool & ef): GenSheetParser[ka1, st, String] =
        let* s1     = content();
        let* ans    = liftAction(_ -> f(s1));
        if (ans) 
            return(s1)
        else 
            throwError("satisfy - failed")


    pub def choice(ps: List[GenSheetParser[ka1, st, a]]): GenSheetParser[ka1, st, a] =
        match ps {
            case p1 :: rs => ifThenElse(p1, return, choice(rs))
            case Nil => fail("choice")
        }

    pub def optional(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Option[a]] = 
        alt(map(Some,p), return(None))

    pub def optionalWithDefault(p: GenSheetParser[ka1, st, a], d: a): GenSheetParser[ka1, st, a] = 
        alt(p, return(d))

    pub def ignore(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Unit] =
        map(constant(), p)

    pub def between(openp: GenSheetParser[ka1, st, open], 
                    closep: GenSheetParser[ka1, st, close], 
                    p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] = 
        let* _ = openp;
        let* a = p;
        let* _ = closep;
        return(a)        

    /// `p` is run with `lookahead`, no text is consumed
    pub def succeeds(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Bool] = 
        alt(seqRight(lookahead(p), return(true)), return(false))

    pub def countWith(combine: (ac, a) -> ac, acc: ac, n: Int32, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, ac] = 
        if (n <= 0)
            return(acc)
        else {
            let* a = p;
            countWith(combine, combine(acc, a), n-1, p)
        }

    pub def count(n: Int32, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        countWith(snoc, empty(), n, p) |> map(toList)


    pub def skipMany(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Unit] = 
        ifThenElse(p, _ -> skipMany(p), return())

    pub def skipMany1(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Unit] = 
        bind(p, _ -> skipMany(p))


    pub def skipManyTill(p: GenSheetParser[ka1, st, a], endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, Unit] = 
        ifThenElse(endp, _ -> return(), bind(p, _ -> skipManyTill(p, endp)))

    pub def skipMany1Till(p: GenSheetParser[ka1, st, a], endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, Unit] = 
        bind(p, _ -> skipManyTill(p, endp))

}

