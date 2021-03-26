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


package flix.runtime.spt.sheetio;

import org.apache.poi.ss.usermodel.CellType;

/* Static wrappers for parts of the POI API that cause trouble for
 * Flix's JVM interface.
 */
public class Wrapper {

    public static CellType get_NONE() {
        return CellType._NONE;
    }

}
