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

mod Internal.Apache.Poi.SS.Usermodel {

    /// Enum

    pub enum CellType(##org.apache.poi.ss.usermodel.CellType)


}

mod Internal.Apache.Poi.SS.Usermodel.CellType {

    use Internal.Apache.Poi.SS.Usermodel.CellType
    use Internal.Apache.Poi.SS.Usermodel.CellType.{CellType}
    use Internal.Apache.Poi.SS.Usermodel.Font
    use Internal.Apache.Poi.SS.Usermodel.Font.{Font}


    pub def ty_none(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType._NONE: ##org.apache.poi.ss.usermodel.CellType \ {} as get_NONE;
        get_NONE() |> CellType

    pub def ty_blank(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType.BLANK: ##org.apache.poi.ss.usermodel.CellType \ {} as get_BLANK;
        get_BLANK() |> CellType

    pub def ty_boolean(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType.BOOLEAN: ##org.apache.poi.ss.usermodel.CellType \ {} as get_BOOLEAN;
        get_BOOLEAN() |> CellType

    pub def ty_error(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType.ERROR: ##org.apache.poi.ss.usermodel.CellType \ {} as get_ERROR;
        get_ERROR() |> CellType

    pub def ty_formula(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType.FORMULA: ##org.apache.poi.ss.usermodel.CellType \ {} as get_FORMULA;
        get_FORMULA() |> CellType

    pub def ty_numeric(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType.NUMERIC: ##org.apache.poi.ss.usermodel.CellType \ {} as get_NUMERIC;
        get_NUMERIC() |> CellType

    pub def ty_string(): CellType = 
        import static get org.apache.poi.ss.usermodel.CellType.STRING: ##org.apache.poi.ss.usermodel.CellType \ {} as get_STRING;
        get_STRING() |> CellType

}
