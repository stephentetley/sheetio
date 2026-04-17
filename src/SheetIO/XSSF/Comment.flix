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


pub mod SheetIO.XSSF.Comment {



    import org.apache.poi.xssf.usermodel.{XSSFComment => JXSSFComment}
    import java.lang.{Exception => JException}

    pub enum XssfComment(JXSSFComment)

    pub eff Comment {
        def getAuthor(): String
    }

    pub def handle(comment: XssfComment, 
                    f: a -> b \ ef): a -> Result[String, b] \ ef - Comment = x -> region rc {
        handleHelper(rc, comment, f, x)
    }

    def handleHelper(_: Region[r],
                        comment: XssfComment, 
                        f: a -> b \ ef): a -> Result[String, b] \ (ef - Comment) + r  = x -> {
        let XssfComment.XssfComment(jcomment) = comment;
        run {
            Ok(f(x))
        } with handler Comment {
            
            def getAuthor(k) = {
                let ans = unsafe IO as r { jcomment.getAuthor() };
                k(ans)
            }
        }
    }

    pub def runComment(comment: XssfComment, f: Unit -> a \ ef): Result[String, a] \ (ef - Comment) = 
        handle(comment, f)()

}