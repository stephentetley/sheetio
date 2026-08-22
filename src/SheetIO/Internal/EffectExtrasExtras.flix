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
// might be consided missing from EffectExtras.

pub mod SheetIO.Internal.EffectExtrasExtras {

    use EffectExtras.Alt
    use EffectExtras.Alt.Alternative
    use EffectExtras.Fail

    pub def runWithFail(f: Unit -> a \ ef): a \ ef - Alternative + Fail = 
        match EffectExtras.Alt.runWithResult(f) {
            case Ok(a)    => a
            case Err(msg) => Fail.report(msg)
        }
}
