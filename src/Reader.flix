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

    pub def readAllRows(sheetName: String, path: System.FilePath, eval: RowEvaluator[a]): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, return, getSheetNamed, allRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(sheetName);
            let* xs     = allRows(s1, eval);
            return(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), path)

    pub def readSomeRows(sheetName: String, path: System.FilePath, eval: RowEvaluator[Option[a]]): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, return, getSheetNamed, someRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(sheetName);
            let* xs     = someRows(s1, eval);
            return(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), path)


    pub def readAllNamedRows(sheetName: String, path: System.FilePath, eval: NamedRowEvaluator[a]): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, return, getSheetNamed, allNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(sheetName);
            let* xs     = allNamedRows(s1, eval);
            return(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), path)

    pub def readSomeNamedRows(sheetName: String, path: System.FilePath, eval: NamedRowEvaluator[Option[a]]): Result[List[a], String] & Impure =
        use SheetIO/SheetReader.{flatMap, return, getSheetNamed, someNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(sheetName);
            let* xs     = someNamedRows(s1, eval);
            return(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), path)

    pub def foldAllRows(sheetName: String, 
                        path: System.FilePath, 
                        combine: (b, a) -> b & ef, 
                        ac: b, 
                        eval: RowEvaluator[a]): Result[b, String] & Impure =
        use SheetIO/SheetReader.{flatMap, return, getSheetNamed, foldAllRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(sheetName);
            let* xs     = foldAllRows(s1, combine, ac, eval);
            return(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), path)


    pub def foldSomeRows(sheetName: String, 
                        path: System.FilePath, 
                        combine: (b, a) -> b & ef, 
                        ac: b, 
                        eval: RowEvaluator[Option[a]]): Result[b, String] & Impure =
        let combine1 = (b, a) -> match a { 
            case Some(a1) => combine(b, a1)
            case None => b as & ef
        };
        foldAllRows(sheetName, path, combine1, ac, eval)


    pub def foldAllNamedRows(sheetName: String, 
                            path: System.FilePath, 
                            combine: (b, a) -> b & ef, 
                            ac: b, 
                            eval: NamedRowEvaluator[a]): Result[b, String] & Impure =
        use SheetIO/SheetReader.{bind, flatMap, return, getSheetNamed, foldAllNamedRows};
        let readStep = () -> { 
            let* s1     = getSheetNamed(sheetName);
            let* xs     = foldAllNamedRows(s1, combine, ac, eval);
            return(xs)
        };
        SheetIO/SheetReader.runXlsx(readStep(), path)

}