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


pub mod SheetIO.Usermodel.Name {




    import org.apache.poi.ss.usermodel.{Name => JName}
    import java.lang.{Exception => JException}

    pub enum Name(JName)


    instance ToJava[Name] {
        type Out = JName
        pub def toJava(x: Name): JName = 
            let Name.Name(x1) = x;
            x1
    }

    instance ToFlix[Name] {
        type In = JName
        pub def toFlix(x: JName): Name = Name.Name(x)
    }

    pub def isHidden(x: Name): Bool =
        let x1 = ToJava.toJava(x);
        unsafe IO {x1.isHidden()}

    pub def getNameName(x: Name): String =
        let x1 = ToJava.toJava(x);
        unsafe IO {x1.getNameName()}

}
