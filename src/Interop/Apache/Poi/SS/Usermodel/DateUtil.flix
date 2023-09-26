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

mod Interop.Apache.Poi.SS.Usermodel.DateUtil {


    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.Cell.{Cell}

    /// Caller should test we have a NUMERIC cell first...
    pub def isCellDateFormatted(cell: Cell): Bool = 
        import static org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(##org.apache.poi.ss.usermodel.Cell): Bool \ {};
        let Cell(cell1) = cell;
        isCellDateFormatted(cell1)

    /// 
    pub def parseDateTime(s: String): Float64 = 
        import static org.apache.poi.ss.usermodel.DateUtil.parseDateTime(String): ##java.lang.Double \ {};
        parseDateTime(s) |> Float64.doubleValue


}

