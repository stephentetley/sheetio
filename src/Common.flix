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



namespace SheetIO {

    /// TODO - do we really need to wrap these?
    
    pub opaque type Sheet = ##org.apache.poi.ss.usermodel.Sheet

    pub opaque type Row = ##org.apache.poi.ss.usermodel.Row

    pub opaque type CellStyle = ##org.apache.poi.ss.usermodel.CellStyle

    pub opaque type DataFormat = ##org.apache.poi.ss.usermodel.DataFormat

    pub opaque type Font = ##org.apache.poi.ss.usermodel.Font

}

