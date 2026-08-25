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


pub mod SheetIO.Usermodel.RichTextString {

    use JavaInterop.AdaptorExtras

    import org.apache.poi.ss.usermodel.{RichTextString => JRichTextString}

    // RichTextString seems like something we might want to first-class
    // and send to and return from functions, hence it is not an effect.
    // We might want a region param and or a type-level boolean if we 
    // want to make it mutable...

    pub enum RichTextString(JRichTextString)

    instance ToString[RichTextString] {
        pub def toString(x: RichTextString): String =
            let RichTextString.RichTextString(x1) = x;
            unsafe IO {x1.getString()}
    }

    instance ToJava[RichTextString] {
        type Out = JRichTextString
        pub def toJava(x: RichTextString): JRichTextString =
            let RichTextString.RichTextString(x1) = x;
            x1
    }

    instance ToFlix[RichTextString] {
        type In = JRichTextString
        pub def toFlix(x: JRichTextString): RichTextString = RichTextString.RichTextString(x)
    }

    pub def length(x: RichTextString): Int32  =
        AdaptorExtras.extract(x, x1 -> unsafe IO {x1.length()})

    pub def numFormattingRuns(x: RichTextString): Int32 =
        AdaptorExtras.extract(x, x1 -> unsafe IO {x1.numFormattingRuns()})

}