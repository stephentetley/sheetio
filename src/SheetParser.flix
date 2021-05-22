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



namespace SheetIO/SheetParser {

    use SheetIO.Sheet;

    type alias SourcePosition = {row: Int32, col: Int32}

    type alias ParseState[st] = {pos: SourcePosition, userState: st}

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



    pub def return(x: a): GenSheetParser[ka1, st, a] = GenSheetParser(sk -> sk(x))    

    pub def throwError(message: String) : GenSheetParser[ka1, st, ans] =
        use Chain.{singleton};
        GenSheetParser((_, fk, _, st, _) -> fk(st.pos, singleton(message)))

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

    pub def liftAction(f: Unit -> a & ef): GenSheetParser[ka1, st, a] =
        GenSheetParser((sk, fk, sheet, st, errs) ->
            let a = f() as & Impure;
            sk(a, fk, sheet, st, errs)
        )
        
    pub def ifThenElse(cond: GenSheetParser[ka1, st, a], 
                        thenp: a -> GenSheetParser[ka1, st, b], 
                        elsep: GenSheetParser[ka1, st, b]): GenSheetParser[ka1, st, b] = 
        GenSheetParser((sk, fk, sheet, st) ->
            apply1(cond, x -> apply1(thenp(x), sk), _ -> apply1(elsep, sk, fk, sheet, st), sheet, st))

}

