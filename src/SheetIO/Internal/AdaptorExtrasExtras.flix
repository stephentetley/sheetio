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

// This is a trial module. The functions here are experiments and
// are considered currently missing from `Interop.Internal.AdaptorExtras`
// in flix-time 


pub mod SheetIO.Internal.AdaptorExtrasExtras {

    use EffectExtras.Fail.Fail

    import java.lang.{Exception => JException}

    pub def guarded(f: Unit -> a \ ef): a \ ef + Fail = 
        try {
            let ans = f();
            Fail.returnNotNull(ans)
        } catch  {
            case ex : JException => Fail.report(ex.getMessage())
        }


}