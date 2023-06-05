/*
 * Copyright 2023 Stephen Tetley
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


mod GradedExtras.Utils {

    use Graded.RgnEnvEvaluator
    use Graded.RgnEnvEvaluator.{RgnEnvEvaluator}



    pub def liftGetterWithAll(f: Region[r] -> env -> vs -> a \ ef): RgnEnvEvaluator[ka, r, env, vs, ef, a]= 
        RgnEnvEvaluator((rc, e, s, k) ->
            try {
                let ans = f(rc, e, s);
                k(ans)
            } catch {
                case ex: ##java.lang.Exception => 
                    import java.lang.Throwable.getMessage(): String \ {};
                    Err(getMessage(ex))
            }
        )

    pub def liftGetterResultWithAll(f: Region[r] -> env -> vs -> Result[err, a] \ ef): RgnEnvEvaluator[ka, r, env, vs, ef, a] with ToString[err] = 
        RgnEnvEvaluator((rc, e, s, k) ->
            match f(rc, e, s) { 
                case Ok(a)    => k(a)
                case Err(err) => Err(ToString.toString(err))
            } 
        )


}
