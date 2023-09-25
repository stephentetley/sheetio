/*
 * Copyright 2022 Stephen Tetley
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


mod Interop.Apache.Poi.SS.Usermodel.WorkbookFactory {

    use Basis.IO.InputStream
    use Basis.IO.InputStream.{InputStream}

    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook
    use Interop.Apache.Poi.SS.Usermodel.MutWorkbook.{MutWorkbook}
    use Interop.Apache.Poi.SS.Usermodel.Workbook
    use Interop.Apache.Poi.SS.Usermodel.Workbook.{Workbook}
    

        
    /// This automatically identifies HSSF or XSSF from the input.
    pub def createWithInputStream(ins: InputStream[r]): Result[String, Workbook] \ r =
        Result.tryCatch(_ -> {
            import static org.apache.poi.ss.usermodel.WorkbookFactory.create(##java.io.InputStream): ##org.apache.poi.ss.usermodel.Workbook \ r;
            let InputStream(ins1) = ins;
            create(ins1) |> Workbook
        })

    /// This automatically identifies HSSF or XSSF from the input.
    pub def createMutWithInputStream(_: Region[r], ins: InputStream[r1]): Result[String, MutWorkbook[r]] \ {r, r1} =
        Result.tryCatch(_ -> {
            import static org.apache.poi.ss.usermodel.WorkbookFactory.create(##java.io.InputStream): ##org.apache.poi.ss.usermodel.Workbook \ {r, r1};
            let InputStream(ins1) = ins;
            create(ins1) |> MutWorkbook
        })

}
