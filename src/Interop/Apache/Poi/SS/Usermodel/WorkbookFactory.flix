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



namespace Interop/Apache/Poi/SS/Usermodel/WorkbookFactory {

    use Interop/Apache/Poi/SS/Usermodel.Workbook;
    use Interop/Apache/Poi/SS/Usermodel.Workbook.{Workbook};
    use Interop/IO.SubInputStream
    
    pub def createWithInputStream![r: Region, ins: Region -> Type]
                                (ins: ins[r]): Result[Workbook[r], String] \ Write(r) with SubInputStream[ins] =
        try {
            import static org.apache.poi.ss.usermodel.WorkbookFactory.create(##java.io.InputStream): ##org.apache.poi.ss.usermodel.Workbook \ r;
            let ins1 = Interop/IO/SubInputStream.toInputStream(ins) |> Interop/IO/InputStream.getHH;
            create(ins1) |> Workbook |> Ok
        } catch {
            case ex: ##java.lang.Exception =>
                import java.lang.Throwable.getMessage(): String \ {};
                Err(getMessage(ex))
        }


}
