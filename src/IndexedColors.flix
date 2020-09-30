/*
 * Copyright 2020 Stephen Tetley
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



namespace SheetIO {

    ///
    /// Represents a IndexedColors enum.
    ///
    pub opaque type IndexedColors = ##org.apache.poi.ss.usermodel.IndexedColors

}

namespace SheetIO/IndexedColors {

    use SheetIO.IndexedColors;
    use SheetIO.IndexedColors.IndexedColors;

    pub def getIndex(x: IndexedColors): Int16 = 
        import org.apache.poi.ss.usermodel.IndexedColors.getIndex();
        let IndexedColors(x1) = x;
        x1.getIndex() as & Pure

    pub def aqua(): IndexedColors = 
        import get org.apache.poi.ss.usermodel.IndexedColors:AQUA as getAQUA ;
        IndexedColors(getAQUA()) as & Pure



}