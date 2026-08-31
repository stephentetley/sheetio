/*
 * Copyright 2026 Stephen Tetley
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

pub mod SheetIO.Common.DateUtil {

    use EffectExtras.Fail.Fail
    use JavaInterop.AdaptorUnchecked
    use InteropTime.LocalDate


    import org.apache.poi.ss.usermodel.{DateUtil => JDateUtil}


    pub def convertTime(s: String): Float64 = 
        unsafe IO { JDateUtil.convertTime(s) }

    pub def isValidExcelDate(d: Float64): Bool = 
        unsafe IO { JDateUtil.isValidExcelDate(d) }


    pub def getExcelDate(d: LocalDate): Float64 = 
        AdaptorUnchecked.extract(d, d1 -> unsafe IO { JDateUtil.getExcelDate(d1) })

    // pub def parseYYYYMMDDDate(s: String): LocalDate = 
    //     unsafe IO { JDateUtil.parseYYYYMMDDDate(s) }


}
 