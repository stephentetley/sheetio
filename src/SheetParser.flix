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

    type alias SourcePosition = (Int32, Int32)

    type alias ParseState[st] = {position: SourcePosition, userState: st}

    type alias Failures = Chain[String]

    type alias FailK[ka1] = SourcePosition -> Failures ~> ka1

    type alias SuccessK[ka1, st, ans] = ans -> FailK[ka1] -> ParseState[st] -> Sheet -> Failures ~> ka1


    /// `st` is the type of user state
    /// `ans` is the type of "immediate answer"
    /// `ka1` is the continuation answer type - if Flix had user visible existentials it would 
    /// be hidden as an existential type. Generally it will always be represented with a type variable
    /// except in "run" functions.
    pub opaque type GenSheetParser[ka1, st, ans] = SuccessK[ka1, st, ans] -> FailK[ka1] -> ParseState[st] -> Sheet -> Failures ~> ka1

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
                let state = {position = (0,0), userState = state};
                let fk = (posn, errs) -> Err({messages = errs, position = posn}) as & Impure;
                let sk = (a, _, _, _, _) -> Ok(a) as & Impure;
                let GenSheetParser(p1) = p;
                let ans = p1(sk, fk, state, sheet, Chain.empty());
                let* _ = SheetIO/Internal/Prim.closeXSSFWorkbook(wb);
                match ans {
                    case Ok(a) => Ok(a)
                    case Err(e1) => Err("Error {row: ${fst(e1.position)}, column: ${snd(e1.position)}}")
                }
            }
        }


    pub def return(x: a): GenSheetParser[ka1, st, a] = GenSheetParser(sk -> sk(x))    

    pub def throwError(message: String) : GenSheetParser[ka1, st, ans] =
        use Chain.{singleton};
        GenSheetParser((_, fk, st, _, _) -> fk(st.position, singleton(message)))

    pub def fail(message: String) : GenSheetParser[ka1, st, ans] = throwError(message)

    def apply1 (p: GenSheetParser[ka1, st, ans], 
                sk: SuccessK[ka1, st, ans], 
                fk: FailK[ka1], 
                st: ParseState[st],
                sheet: Sheet,  
                errs: Failures): ka1 & Impure =
        let GenSheetParser(pf) = p;
        pf(sk, fk, st, sheet, errs)


    ///
    /// `swapError` is quite subtle. Swapping the error continuation needs to be
    /// a "local" operation which is undone (restored) afterwards (c.f. local in the Reader
    /// monad). It must not be a "stateful" operation with unlimited extent.
    ///
    pub def swapError(p: GenSheetParser[ka1, st, ans], message: String) : GenSheetParser[ka1, st, ans] =
        use Chain.{singleton};
        GenSheetParser((sk, fk, st, sheet, errs) ->
            let fk1 = (pos, _) -> fk(pos, singleton(message));
            let sk1 = (x, _, st1, _, _) -> sk(x, fk, st1, sheet, errs);
            apply1(p, sk1, fk1, st, sheet, errs)
        )


    pub def map(f: a -> b & ef, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, b] =
        GenSheetParser((sk, fk, st, sheet, errs) ->
            let sk1 = (x, fk1, st1, _, errs1) -> sk(f(x) as & Impure, fk1, st1, sheet, errs1);
            apply1(p, sk1, fk, st, sheet, errs)
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


    pub def ifThenElse(cond: GenSheetParser[ka1, st, a], 
                        thenp: a -> GenSheetParser[ka1, st, b], 
                        elsep: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, b] = 
        GenSheetParser((sk, fk, st, sheet) ->
            apply1(cond, (x, fk1, st1, _) -> apply1(thenp(x), sk, fk1, st1, sheet), _ -> apply1(elsep, sk, fk, st, sheet), st, sheet))


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
        GenSheetParser((sk, fk, st, sheet, errs) ->
            let a = f() as & Impure;
            sk(a, fk, st, sheet, errs)
        )
        
    pub def liftSheetPrim(f: Sheet ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        use Chain.{singleton};
        GenSheetParser((sk, fk, st, sheet, errs) ->
            let ans = f(sheet);
            match ans { 
                case Ok(a) => sk(a, fk, st, sheet, errs)
                case Err(msg) => fk(st.position, singleton("liftSheetPrim - ${msg}"))
            }
        )

    pub def liftRowPrim(f: ##org.apache.poi.ss.usermodel.Row ~> Result[a, String]): GenSheetParser[ka1, st, a] =
        getRowNumber() >>= ix -> liftRowPrimHelper(ix, f)

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
        getColumnNumber() >>= ix -> liftCellPrimHelper(ix, f)

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
    
    def getRowNumber(): GenSheetParser[ka1, st, Int32] =  
        map(fst, getSourcePosition())
    
    def getColumnNumber(): GenSheetParser[ka1, st, Int32] =  
        map(snd, getSourcePosition())

    pub def getSourcePosition(): GenSheetParser[ka1, st, SourcePosition] =  
        GenSheetParser((sk, fk, st, sheet, errs) ->
            let pos = st.position;
            sk(pos, fk, st, sheet, errs)
        )

    
    pub def setSourcePosition(row: Int32, col: Int32): GenSheetParser[ka1, st, Unit] =  
        GenSheetParser((sk, fk, st, sheet) ->
            let st1 = {position = (row, col) | st};
            sk((), fk, st1, sheet)
        ) 

    pub def gotoRow(row: Int32): GenSheetParser[ka1, st, Unit] = setSourcePosition(row, 0)

    pub def getState(): GenSheetParser[ka1, st, st] =  
        GenSheetParser((sk, fk, st, sheet) ->            
            sk(st.userState, fk, st, sheet)
        )

    pub def setState(state: st): GenSheetParser[ka1, st, Unit] =  
        GenSheetParser((sk, fk, st, sheet) ->
            let st1 = {userState = state | st};
            sk((), fk, st1, sheet)
        )


    pub def updateState(modify: st -> st): GenSheetParser[ka1, st, Unit] =  
        GenSheetParser((sk, fk, st, sheet) ->
            let st1 = {userState = modify(st.userState) | st};
            sk((), fk, st1, sheet)
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
        let* xy = getSourcePosition();
        let (row, col) = xy;
        setSourcePosition(row, col + 1)
        
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

    pub def localDateTime(): GenSheetParser[ka1, st, Time.LocalDateTime] =
        use Time.LocalDateTime.{LocalDateTime};
        map(LocalDateTime, readCellWith(SheetIO/Internal/Cell.getLocalDateTimeCellValue))

    pub def localDate(): GenSheetParser[ka1, st, Time.LocalDate] =
        map(Time/LocalDateTime.toLocalDate, localDateTime())



    pub def literal(s: String): GenSheetParser[ka1, st, String] =
        let* s1 = content();
        if (s == s1) 
            return(s)
        else 
            throwError("literal - no match")

    pub def row(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =  
        let* ans        = p;
        let* row        = getRowNumber();
        let* _          = setSourcePosition(row + 1, 0);
        return(ans)

    pub def skipRow(): GenSheetParser[ka1, st, Unit] =  
        let* end        = isAtEnd();
        if (end) 
            throwError("skipRow - at End")
        else {
            let* row        = getRowNumber();
            let* _          = setSourcePosition(row + 1, 0);
            return()
        }

    ///
    /// lookahead - run the parser but don't consume input on success
    /// On fail, fail.
    ///
    pub def lookahead(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =
        GenSheetParser((sk, fk, st, sheet) ->
            let sk1 = (ans, fk1, _, _) -> sk(ans, fk1, st, sheet);
            apply1(p, sk1, fk, st, sheet)
        )

    pub def endOfInput() : GenSheetParser[ka1, st, Unit] =
        GenSheetParser((sk, fk, st, sheet, errs) -> 
            match SheetIO/Internal/Sheet.getLastRowNum(sheet) {
                case Err(msg) => fk(st.position, Chain.cons("endOfInput - ${msg}", errs))
                case Ok(rmax) => 
                    if (fst(st.position) >= rmax) 
                        sk((), fk, st, sheet, errs)
                    else 
                        fk(st.position, Chain.cons("endOfInput - not at end", errs))
            }
        )

    pub def isAtEnd() : GenSheetParser[ka1, st, Bool] =
        GenSheetParser((sk, fk, st, sheet, errs) -> 
            match SheetIO/Internal/Sheet.getLastRowNum(sheet) {
                case Err(msg) => fk(st.position, Chain.cons("isEof - ${msg}", errs))
                case Ok(rmax) => 
                    if (fst(st.position) >= rmax) 
                        sk(true, fk, st, sheet, errs)
                    else 
                        sk(false, fk, st, sheet, errs)
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
        alt(map(Some, p), return(None))

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


    pub def manyWith(combine: (ac, a) -> ac, acc: ac, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, ac] = 
        ifThenElse(p, a -> manyWith(combine, combine(acc, a), p), return(acc))


    pub def many(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        manyWith(snoc, empty(), p) |> map(toList)
        
    pub def many1With(combine: (ac, a) -> ac, acc: ac, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, ac] = 
        bind(p, x -> manyWith(combine, combine(acc, x), p))
        
    pub def many1(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        many1With(snoc, empty(), p) |> map(toList)


    /// Count how many times parser `p` succeeds
    pub def ntimes(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Int32] = 
        manyWith((ac, _) -> ac + 1, 0, p)


    pub def manyTillWith(combine: (ac, a) -> ac, 
                            acc: ac, 
                            p: GenSheetParser[ka1, st, a], 
                            endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, ac] = 
        ifThenElse(endp, _ -> return(acc), bind(p, a -> manyTillWith(combine, combine(acc, a), p, endp)))

    pub def manyTill(p: GenSheetParser[ka1, st, a], endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        manyTillWith(snoc, empty(), p, endp) |> map(toList)
        
    pub def many1TillWith(combine: (ac, a) -> ac, 
                            acc: ac, 
                            p: GenSheetParser[ka1, st, a], 
                            endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, ac] = 
        bind(p, x -> manyTillWith(combine, combine(acc, x), p, endp))
        
    pub def many1Till(p: GenSheetParser[ka1, st, a], endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        many1TillWith(snoc, empty(), p, endp) |> map(toList)
    

    pub def sepByWith(combine: (ac, a) -> ac, 
                        acc: ac, 
                        p: GenSheetParser[ka1, st, a], 
                        sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        ifThenElse(p, x -> sepByWithHelper(combine, combine(acc, x), p, sep), return(acc))
    
    // Initial p has been consumed
    def sepByWithHelper(combine: (ac, a) -> ac, 
                        acc: ac, 
                        p: GenSheetParser[ka1, st, a], 
                        sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        ifThenElse(seqRight(sep, p), x -> sepByWithHelper(combine, combine(acc, x), p, sep), return(acc))

    pub def sepBy(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        sepByWith(snoc, empty(), p, sep) |> map(toList)

    pub def sepBy1With(combine: (ac, a) -> ac, 
                        acc: ac, 
                        p: GenSheetParser[ka1, st, a], 
                        sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        bind(p, x -> sepByWithHelper(combine, combine(acc, x), p, sep))
    
    pub def sepBy1(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        sepBy1With(snoc, empty(), p, sep) |> map(toList)

    pub def endByWith(combine: (ac, a) -> ac, 
                        acc: ac, 
                        p: GenSheetParser[ka1, st, a], 
                        sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        manyWith(combine, acc, seqLeft(p, sep))

    pub def endBy(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        endByWith(snoc, empty(), p, sep) |> map(toList)

    pub def endBy1With(combine: (ac, a) -> ac, 
                        acc: ac, 
                        p: GenSheetParser[ka1, st, a], 
                        sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        bind(seqLeft(p, sep), x -> endByWith(combine, combine(acc, x), p, sep))

    pub def endBy1(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        endBy1With(snoc, empty(), p, sep) |> map(toList)

    // seperated and optionally terminated with sep.
    pub def sepEndByWith(combine: (ac, a) -> ac, 
                            acc: ac, 
                            p: GenSheetParser[ka1, st, a], 
                            sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        let* xs = sepByWith(combine, acc, p, sep);
        ifThenElse(sep, _ -> return(xs), return(xs))

    // seperated and optionally terminated with sep.
    pub def sepEndBy(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        map(toList, sepEndByWith(snoc, empty(), p, sep))

    // seperated and optionally terminated with sep.
    pub def sepEndBy1With(combine: (ac, a) -> ac, 
                            acc: ac, 
                            p: GenSheetParser[ka1, st, a], 
                            sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, ac] = 
        let* xs = sepBy1With(combine, acc, p, sep);
        ifThenElse(sep, _ -> return(xs), return(xs))

    // seperated and optionally terminated with sep.
    pub def sepEndBy1(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        sepEndBy1With(snoc, empty(), p, sep) |> map(toList)

    /// Count how many times parser `p` succeeds
    pub def ntimesSepBy(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, Int32] = 
        sepByWith((ac, _) -> ac + 1, 0, p, sep)

    /// Count how many times parser `p` succeeds
    pub def ntimesEndBy(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, Int32] = 
        endByWith((ac, _) -> ac + 1, 0, p, sep)

    /// Count how many times parser `p` succeeds
    pub def ntimesSepEndBy(p: GenSheetParser[ka1, st, a], sep: GenSheetParser[ka1, st, sep]): GenSheetParser[ka1, st, Int32] = 
        sepEndByWith((ac, _) -> ac + 1, 0, p, sep)

    /// p1 must succeed, p2 parsers zero-or-more
    pub def manyTill2With(combine: (ac, a) -> ac, 
                            acc: ac, 
                            p1: GenSheetParser[ka1, st, a], 
                            p2: GenSheetParser[ka1, st, a], 
                            endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, ac] = 
        bind(p1, x -> manyTillWith(combine, combine(acc, x), p2, endp))
        
    /// p1 must succeed, p2 parsers zero-or-more
    pub def manyTill2(p1: GenSheetParser[ka1, st, a], 
                      p2: GenSheetParser[ka1, st, a], 
                      endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        manyTill2With(snoc, empty(), p1, p2, endp) |> map(toList)

    /// p1 must succeed, p2 parsers one-or-more
    pub def many1Till2With(combine: (ac, a) -> ac, acc: ac, 
                            p1: GenSheetParser[ka1, st, a], 
                            p2: GenSheetParser[ka1, st, a], 
                            endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, ac] = 
        bind(p1, x -> many1TillWith(combine, combine(acc, x), p2, endp))
        
    /// p1 must succeed, p2 parsers one-or-more
    pub def many1Till2(p1: GenSheetParser[ka1, st, a], 
                        p2: GenSheetParser[ka1, st, a], 
                        endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, List[a]] = 
        use Chain.{empty, snoc, toList};
        many1Till2With(snoc, empty(), p1, p2, endp) |> map(toList)

    pub def skipMany(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Unit] = 
        ifThenElse(p, _ -> skipMany(p), return())

    pub def skipMany1(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, Unit] = 
        bind(p, _ -> skipMany(p))

    /// This works fine vis-a-vis not breaking the stack in FlixParsec...
    pub def skipManyTill(p: GenSheetParser[ka1, st, a], endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, Unit] = 
        ifThenElse(endp, _ -> return(), bind(p, _ -> skipManyTill(p, endp)))


    pub def skipMany1Till(p: GenSheetParser[ka1, st, a], endp: GenSheetParser[ka1, st, end]): GenSheetParser[ka1, st, Unit] = 
        bind(p, _ -> skipManyTill(p, endp))

    /// Note - bounded and find don't play well together with the current implementation....
    /// Bounded won't curtial the extent of find.
    
    /// TODO - allow find to wrap around...
    pub def find(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] = 
        ifThenElse(p, return, bind(skipRow(), _ -> find(p)))

    
    pub def island(p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] = lookahead(find(p))


    ///
    /// Run parser `p` within a bounded number of rows `nrowsWindow`.
    ///
    pub def bounded(nrowsWindow: Int32, p: GenSheetParser[ka1, st, a]): GenSheetParser[ka1, st, a] =
        if (nrowsWindow > 0)
            ifThenElse(p, return, bind(skipRow(), _ -> bounded(nrowsWindow - 1, p)))
        else
            throwError("bounded - out-of-bounds")

    pub def chainl1(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a]): GenSheetParser[ka1, st, a] = 
        bind(p, v1 -> chainl1Helper(p, op, v1))
        
    def chainl1Helper(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a], x: a): GenSheetParser[ka1, st, a] = 
        alt(bind(op, f -> bind(p, v1 -> chainl1Helper(p, op, f(x, v1)))), 
            return(x))

    pub def chainl(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a], x: a): GenSheetParser[ka1, st, a] = 
        alt(chainl1(p,op), return(x))


    def chainr1Helper(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a], x: a): GenSheetParser[ka1, st, a] = 
        alt(chainr1Step(p, op, x), return(x))

    def chainr1Step(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a], x: a): GenSheetParser[ka1, st, a] = 
        bind(op, f -> 
            bind(bind(p, v1 -> chainr1Helper(p, op, v1)), y -> 
                bind(chainr1Helper(p,op, f(x,y)), ans -> return(ans))))

    pub def chainr1(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a]): GenSheetParser[ka1, st, a] = 
        bind(p, v1 -> chainr1Helper(p, op, v1))

    pub def chainr(p: GenSheetParser[ka1, st, a], op: GenSheetParser[ka1, st, (a, a) -> a], x: a): GenSheetParser[ka1, st, a] = 
        alt(chainr1(p, op), return(x))


}

