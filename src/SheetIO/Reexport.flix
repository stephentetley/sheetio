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

    pub type alias Cell[r: Region] = Interop.Apache.Poi.SS.Usermodel.Cell[r]

    pub type alias CellStyle[r: Region] = Interop.Apache.Poi.SS.Usermodel.CellStyle[r]

    pub type alias DataFormat[r: Region] = Interop.Apache.Poi.SS.Usermodel.DataFormat[r]

    pub type alias Font[r: Region] = Interop.Apache.Poi.SS.Usermodel.Font[r]

    pub type alias Row[r: Region] = Interop.Apache.Poi.SS.Usermodel.Row[r]
    
    pub type alias Sheet[r: Region] = Interop.Apache.Poi.SS.Usermodel.Sheet[r]

    pub type alias Workbook[r: Region] = Interop.Apache.Poi.SS.Usermodel.Workbook[r]

}
