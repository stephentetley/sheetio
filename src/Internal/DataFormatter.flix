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

namespace SheetIO/Internal/DataFormatter {


    pub def new(): ##org.apache.poi.ss.usermodel.DataFormatter & Impure = 
        import new org.apache.poi.ss.usermodel.DataFormatter(): ##org.apache.poi.ss.usermodel.DataFormatter & Impure as newDataFormatter;
        newDataFormatter()

    /// TODO - it would be useful to allow users to make custom DataFormatters...

    pub def formatCellValue(formatter: ##org.apache.poi.ss.usermodel.DataFormatter,
                            cell: ##org.apache.poi.ss.usermodel.Cell): Result[String, String] & Impure = 
        import org.apache.poi.ss.usermodel.DataFormatter.formatCellValue(##org.apache.poi.ss.usermodel.Cell): String & Impure; 
        try {
            let ans = formatCellValue(formatter, cell);
            Ok(ans)
        } catch {
            case _: ##java.lang.Exception => Err("formatCellValue")
        }

}

