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


pub mod SheetIO.Common.WorkbookUtil {


    import org.apache.poi.ss.util.{WorkbookUtil => JWorkbookUtil}
    import java.lang.{Exception => JException}
    import java.lang.{IllegalArgumentException => JIllegalArgumentException}


    pub def validateSheetName(s: String): Bool = 
        try {
            let _ = unsafe IO { JWorkbookUtil.validateSheetName(s) };
            true
        } catch {
            case _: JIllegalArgumentException => false
        }


    pub def createSafeSheetName(s: String): Result[String, String] = 
        try {
            let ans = unsafe IO { JWorkbookUtil.createSafeSheetName(s) };
            Ok(ans)
        } catch {
            case _: JIllegalArgumentException => Err("")
        }


}
