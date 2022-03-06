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



namespace SheetIO/Reader {

    use SheetIO.NamedRowEvaluator;
    use SheetIO.RowEvaluator;

    pub type alias XlsxSource = 
        { sheet :: String
        , path :: System.FilePath
        }


    pub def allRows(eval: RowEvaluator[a], src: XlsxSource): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = allRows(eval, s1);
            point(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), src.path)

    pub def someRows(eval: RowEvaluator[Option[a]], src: XlsxSource): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, point, getSheetNamed, someRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = someRows(eval, s1);
            point(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), src.path)


    pub def allNamedRows(eval: NamedRowEvaluator[a], src: XlsxSource): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, point, getSheetNamed, allNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = allNamedRows(eval, s1);
            point(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), src.path)

    pub def someNamedRows(eval: NamedRowEvaluator[Option[a]], src: XlsxSource): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, point, getSheetNamed, someNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = someNamedRows(eval, s1);
            point(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), src.path)

    pub def foldAllRows(combine: (b, a) -> b & ef, 
                        ac: b, 
                        eval: RowEvaluator[a], src: XlsxSource): Result[b, String] & Impure =
        use SheetIO/SheetReader.{flatMap, point, getSheetNamed, foldAllRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = foldAllRows(combine, ac, eval, s1);
            point(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), src.path)


    pub def foldSomeRows(combine: (b, a) -> b & ef, 
                            ac: b, 
                            eval: RowEvaluator[Option[a]], 
                            src: XlsxSource): Result[b, String] & Impure =
        let combine1 = (b, a) -> match a { 
            case Some(a1) => combine(b, a1)
            case None => b as & ef
        };
        foldAllRows(combine1, ac, eval, src)


    pub def foldAllNamedRows(combine: (b, a) -> b & ef, 
                                ac: b, 
                                eval: NamedRowEvaluator[a],
                                src: XlsxSource): Result[b, String] & Impure =
        use SheetIO/SheetReader.{bind, flatMap, point, getSheetNamed, foldAllNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = foldAllNamedRows(combine, ac, eval, s1);
            point(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), src.path)


    pub def foldSomeNamedRows(combine: (b, a) -> b & ef, 
                            ac: b, 
                            eval: NamedRowEvaluator[Option[a]],
                            src: XlsxSource): Result[b, String] & Impure =
        let combine1 = (b, a) -> match a { 
            case Some(a1) => combine(b, a1)
            case None => b as & ef
        };
        foldAllNamedRows(combine1, ac, eval, src)

    pub def allRelations(eval: RowEvaluator[#{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllRows((ac, x) -> ac <+> x, #{}, eval, src)


    pub def allNamedRelations(eval: NamedRowEvaluator[#{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllNamedRows((ac, x) -> ac <+> x, #{}, eval, src)

    pub def someRelations(eval: RowEvaluator[#{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllRows((ac, x) -> ac <+> x, #{}, SheetIO/RowEvaluator.withDefault(eval, #{}), src)


    pub def someNamedRelations(eval: NamedRowEvaluator[#{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllNamedRows((ac, x) -> ac <+> x, #{}, SheetIO/NamedRowEvaluator.withDefault(eval, #{}), src)

}
