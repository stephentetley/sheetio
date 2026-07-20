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


pub mod SheetIO.Usermodel.ReadComment {


    use SheetIO.Usermodel.{Comment}

    import org.apache.poi.ss.usermodel.{Comment => JComment}
    import java.lang.{Exception => JException}

    // TODO what about the actual string? RichTestString might 
    // not be work an effect...

    pub eff ReadComment {
        def getAuthor(): String
        def isVisible(): Bool
    }

    

    pub def handle(comment: Comment, 
                    f: a -> b \ ef): a -> Result[String, b] \ ef - ReadComment = x -> {
        let jcomment = ToJava.toJava(comment);
        run {
            Ok(f(x))
        } with handler ReadComment {
            
            def getAuthor(k) = {
                let ans = unsafe IO as {ef - ReadComment} { jcomment.getAuthor() };
                k(ans)
            }
            
            def isVisible(k) = {
                let ans = unsafe IO as {ef - ReadComment} { jcomment.isVisible() };
                k(ans)
            }

        }
    }

    pub def runComment(comment: Comment, f: Unit -> a \ ef): Result[String, a] \ (ef - ReadComment) = 
        handle(comment, f)()

}