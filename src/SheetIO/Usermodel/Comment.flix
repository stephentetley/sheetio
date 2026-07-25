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


pub mod SheetIO.Usermodel.Comment {

    use SheetIO.Usermodel.RichTextString

    import org.apache.poi.ss.usermodel.{Comment => JComment}
    import java.lang.{Exception => JException}

    pub enum Comment(JComment)

    instance ToString[Comment] {
        pub def toString(x: Comment): String =
            let Comment.Comment(x1) = x;
            unsafe IO {x1.getString().getString()}
    }


    instance ToJava[Comment] {
        type Out = JComment
        pub def toJava(x: Comment): JComment = 
            let Comment.Comment(x1) = x;
            x1
    }

    instance ToFlix[Comment] {
        type In = JComment
        pub def toFlix(x: JComment): Comment = Comment.Comment(x)
    }

    pub def isVisible(x: Comment): Bool =
        let x1 = ToJava.toJava(x);
        unsafe IO {x1.isVisible()}

    pub def getString(x: Comment): RichTextString =
        let x1 = ToJava.toJava(x);
        let ans = unsafe IO {x1.getString()};
        ToFlix.toFlix(ans)

}
