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


namespace MonadTemp/MonadExtra {

    use Applicative.{point};
    use Monad.{>>=};


    pub def whenTrue(x: Bool, ma: m[Unit]): m[Unit] with Monad[m] = 
        match x { 
            case true => ma
            case false => point()
        }

    pub def whenFalse(x: Bool, ma: m[Unit]): m[Unit] with Monad[m] = 
        match x { 
            case true => point()
            case false => ma
        }

    pub def whenTrueM(mx: m[Bool], ma: m[Unit]): m[Unit] with Monad[m] =
        mx >>= x -> whenTrue(x, ma)

    pub def whenFalseM(mx: m[Bool], ma: m[Unit]): m[Unit] with Monad[m] = 
        mx >>= x -> whenFalse(x, ma)


    pub def whenSome(x: Option[a], f: a -> m[Unit] & ef):  m[Unit] & ef with Monad[m] = 
        match x { 
            case Some(a) => f(a)
            case None => point()
        }

    pub def whenNone(x: Option[a], ma:  m[Unit]):  m[Unit] with Monad[m] = 
        match x { 
            case Some(_) => point()
            case None => ma
        }

    pub def whenSomeM(mx:  m[Option[a]], f: a ->  m[Unit] & ef):  m[Unit] & ef with Monad[m] = 
        mx >>= x -> whenSome(x, f)
            

    pub def whenNoneM(mx:  m[Option[a]], ma:  m[Unit]):  m[Unit] with Monad[m] = 
        mx >>= x -> whenNone(x, ma)


    pub def ifSomeM(mx: m[Option[a]], mthen: a -> m[b] & ef, melse: m[b]): m[b] & ef with Monad[m] = 
        mx >>= x -> match x { 
            case Some(a) => mthen(a)
            case None    => melse
        }
            

    pub def ifNoneM(mx: m[Option[a]], mthen: m[b], melse: m[b]): m[b] with Monad[m] = 
        mx >>= x -> match x { 
            case Some(_) => mthen
            case None    => melse
        }


/*
    pub def mapM(f: a -> m[b], t: t[a]): m[t[b]] with Monad[m], Traversable[t] = 
        Traversable.traverse(f, t)

    pub def mapM_(f: a -> m[b], t: t[a]): m[Unit] with Monad[m], Foldable[t] = 
        Foldable.foreachM(f, t)

    pub def sequence(t: t[m[a]]): m[t[a]] with Monad[m], Traversable[t] = 
        Traversable.sequence(t)
*/

}
