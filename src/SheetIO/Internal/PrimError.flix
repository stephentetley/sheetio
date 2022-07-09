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

namespace SheetIO/Internal/PrimError {



    pub type alias PrimError = 
        { source :: String
        , exnMessage :: String
        }

    pub def primError(_: Region[r], source: String, ex: ##java.lang.Exception): PrimError \ Read(r) = 
        import java.lang.Throwable.getMessage(): String & r;
        {source = source, exnMessage = getMessage(ex as ##java.lang.Throwable)} 

    pub def primErr(_: Region[r], source: String, ex: ##java.lang.Exception): Result[a, PrimError] \ Read(r) = 
        import java.lang.Throwable.getMessage(): String & r;
        Err({source = source, exnMessage = getMessage(ex as ##java.lang.Throwable)})

    pub def getText(err: {source :: String, exnMessage :: String | a}): String = 
        if (String.isEmpty(err.exnMessage)) 
            err.source
        else 
            "${err.source}: ${err.exnMessage}"


}

