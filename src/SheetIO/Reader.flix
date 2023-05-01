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



mod SheetIO.Reader {

    use GradedFunctor.{<$>}

    use SheetIO.NamedRowEvaluator
    use SheetIO.RowEvaluator
    use SheetIO.SheetReader
/*
    pub type alias XlsxSource = 
        { sheet :: String
        , path :: Interop.System.Path
        }


    pub def allRows(eval: RowEvaluator[r, a], src: XlsxSource): Result[List[a], String] & Impure =
        SheetIO.SheetReader.runXlsx(allRows1(eval, src.sheet), src.path)

    pub def allRowsOfSheets(eval: RowEvaluator[r, a], sheets: List[String], path: FlixSandbox.System.Path): Result[List[a], String] & Impure =
        let readSheets = List.flatten <$> Traversable.traverse(allRows1(eval), sheets);        
        SheetIO.SheetReader.runXlsx(readSheets, path)


    def allRows1(eval: RowEvaluator[r, a], sheetName: String): SheetReader[r, List[a]] =
        use SheetIO.SheetReader.{flatMap, point, getSheetNamed, allRows};
        let* s1     = getSheetNamed(sheetName);
        let* xs     = allRows(eval, s1);
        point(xs)

    pub def someRows(eval: RowEvaluator[r, Option[a]], src: XlsxSource): Result[List[a], String] & Impure =        
        SheetIO.SheetReader.runXlsx(someRows1(eval, src.sheet), src.path)

    pub def someRowsOfSheets(eval: RowEvaluator[r, Option[a]], sheets: List[String], path: FlixSandbox.System.Path): Result[List[a], String] & Impure =
        let readSheets = List.flatten <$> Traversable.traverse(someRows1(eval), sheets);        
        SheetIO.SheetReader.runXlsx(readSheets, path)


    def someRows1(eval: RowEvaluator[r, Option[a]], sheetName: String): SheetReader[r, List[a]] =
        use SheetIO.SheetReader.{flatMap, point, getSheetNamed, someRows};
        let* s1     = getSheetNamed(sheetName);
        let* xs     = someRows(eval, s1);
        point(xs)


    pub def allNamedRows(eval: NamedRowEvaluator[r, a], src: XlsxSource): Result[List[a], String] & Impure =
        SheetIO.SheetReader.runXlsx(allNamedRows1(eval, src.sheet), src.path)

    pub def allNamedRowsOfSheets(eval: NamedRowEvaluator[r, a], sheets: List[String], path: FlixSandbox.System.Path): Result[List[a], String] & Impure =
        let readSheets = List.flatten <$> Traversable.traverse(allNamedRows1(eval), sheets);        
        SheetIO.SheetReader.runXlsx(readSheets, path)


    def allNamedRows1(eval: NamedRowEvaluator[r, a], sheetName: String): SheetReader[r, List[a]] =
        use SheetIO.SheetReader.{flatMap, point, getSheetNamed, allNamedRows};
        let* s1     = getSheetNamed(sheetName);
        let* xs     = allNamedRows(eval, s1);
        point(xs)
        
    

    pub def someNamedRows(eval: NamedRowEvaluator[r, Option[a]], src: XlsxSource): Result[List[a], String] & Impure =
        SheetIO.SheetReader.runXlsx(someNamedRows1(eval, src.sheet), src.path)

    pub def someNamedRowsOfSheets(eval: NamedRowEvaluator[r, Option[a]], sheets: List[String], path: FlixSandbox.System.Path): Result[List[a], String] & Impure =
        let readSheets = List.flatten <$> Traversable.traverse(someNamedRows1(eval), sheets);        
        SheetIO.SheetReader.runXlsx(readSheets, path)


    def someNamedRows1(eval: NamedRowEvaluator[r, Option[a]], sheetName: String): SheetReader[r, List[a]] =
        use SheetIO.SheetReader.{flatMap, point, getSheetNamed, someNamedRows};
        let* s1     = getSheetNamed(sheetName);
        let* xs     = someNamedRows(eval, s1);
        point(xs)
    

    pub def foldAllRows(combine: (b, a) -> b & ef, 
                        ac: b, 
                        eval: RowEvaluator[r, a], src: XlsxSource): Result[b, String] & Impure =
        use SheetIO.SheetReader.{flatMap, point, getSheetNamed, foldAllRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = foldAllRows(combine, ac, eval, s1);
            point(xs)
        };
        SheetIO.SheetReader.runXlsx(readStep(), src.path)


    pub def foldSomeRows(combine: (b, a) -> b & ef, 
                            ac: b, 
                            eval: RowEvaluator[r, Option[a]], 
                            src: XlsxSource): Result[b, String] & Impure =
        let combine1 = (b, a) -> match a { 
            case Some(a1) => combine(b, a1)
            case None => b as & ef
        };
        foldAllRows(combine1, ac, eval, src)


    pub def foldAllNamedRows(combine: (b, a) -> b & ef, 
                                ac: b, 
                                eval: NamedRowEvaluator[r, a],
                                src: XlsxSource): Result[b, String] & Impure =
        use SheetIO.SheetReader.{bind, flatMap, point, getSheetNamed, foldAllNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(src.sheet);
            let* xs     = foldAllNamedRows(combine, ac, eval, s1);
            point(xs)
        };
        SheetIO.SheetReader.runXlsx(readStep(), src.path)


    pub def foldSomeNamedRows(combine: (b, a) -> b & ef, 
                            ac: b, 
                            eval: NamedRowEvaluator[r, Option[a]],
                            src: XlsxSource): Result[b, String] & Impure =
        let combine1 = (b, a) -> match a { 
            case Some(a1) => combine(b, a1)
            case None => b as & ef
        };
        foldAllNamedRows(combine1, ac, eval, src)

    pub def allRelations(eval: RowEvaluator[r, #{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllRows((ac, x) -> ac <+> x, #{}, eval, src)


    pub def allNamedRelations(eval: NamedRowEvaluator[r, #{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllNamedRows((ac, x) -> ac <+> x, #{}, eval, src)

    pub def someRelations(eval: RowEvaluator[r, #{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllRows((ac, x) -> ac <+> x, #{}, SheetIO.RowEvaluator.withDefault(eval, #{}), src)


    pub def someNamedRelations(eval: NamedRowEvaluator[r, #{| reln}], src: XlsxSource): Result[#{| reln}, String] & Impure =
        foldAllNamedRows((ac, x) -> ac <+> x, #{}, SheetIO.NamedRowEvaluator.withDefault(eval, #{}), src)

*/
}
