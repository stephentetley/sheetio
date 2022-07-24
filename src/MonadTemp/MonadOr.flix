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

class MonadOr[m: Type -> Type] with Monad[m] {

    ///
    /// "Left catch" / left biased choice
    ///
    pub def orelse(a: m[a], b: m[a]): m[a]

}

namespace MonadOr {

    use Functor.{<$>};
    use Applicative.{point};

    pub def <|>(a: m[a], b: m[a]): m[a] with MonadOr[m] = orelse(a, b)

    pub def optional(ma: m[a]): m[Option[a]] with MonadOr[m] = 
        orelse(Some <$> ma, point(None))

    pub def withDefault(d: a, x: m[a]): m[a] with MonadOr[m] =
        x <|> point(d)

}


