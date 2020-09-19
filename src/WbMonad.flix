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

use System.Error.{Generic};

use SheetReader.WbMonad;
use SheetReader.WbMonad.{WbMonad};

namespace SheetReader {

    pub opaque type WbMonad[a] = ##org.apache.poi.xssf.usermodel.XSSFWorkbook ~> Result[a, System.Error]

}

namespace SheetReader/WbMonad {

    pub def runWbMonad(ma: WbMonad[a], path: System.FilePath): Result[a, System.Error] & Impure = 
        use SheetReader/Internal/Prim.{newXSSFWorkbook, closeXSSFWorkbook};
        use Result.{flatMap};
        let path1 = System/FilePath.toString(path);
        match newXSSFWorkbook(path1) { 
            case Err(e) => Err(e)
            case Ok(wb) => {
                let ans = apply1(ma, wb);
                let* _ = closeXSSFWorkbook(wb);
                ans         
            }
        }
        
    def apply1(ma: WbMonad[a], wb: ##org.apache.poi.xssf.usermodel.XSSFWorkbook): Result[a, System.Error] & Impure = 
        let WbMonad(x1) = ma;
        x1(wb)

    pub def return(x: a): WbMonad[a] = 
        WbMonad(_ -> Ok(x) as & Impure)

    pub def throwError(msg: String): WbMonad[a] =
        WbMonad(_ -> Err(System/Error.Generic(msg)) as & Impure)


    pub def catchError(ma: WbMonad[a], handle: System.Error -> WbMonad[a]): WbMonad[a] =
        WbMonad(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(e) => apply1(handle(e), wb)
        })

    pub def mapError(f: System.Error -> System.Error, ma: WbMonad[a]): WbMonad[a] =
        WbMonad(wb -> match apply1(ma, wb) { 
            case Ok(a) => Ok(a)
            case Err(e) => Err(f(e))
        })

    /// Needs to allow allow effects (e.g create / access arrays)
    pub def map(fn: a -> b & ef, ma: WbMonad[a]): WbMonad[b] = 
        WbMonad(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => { let ans = fn(a) as & Impure; Ok(ans) }
                case Err(e) => Err(e)
            }
        )

    pub def ap(mf: WbMonad[a -> b & ef], ma: WbMonad[a]): WbMonad[b] = 
        WbMonad(wb -> 
            match apply1(mf, wb) { 
                case Ok(f) => match apply1(ma, wb) { 
                    case Ok(a) => { let ans = f(a) as & Impure; Ok(ans) }
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqLeft(ma: WbMonad[a], mb: WbMonad[b]): WbMonad[a] = 
        WbMonad(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => match apply1(mb, wb) {
                    case Ok(_) => Ok(a)
                    case Err(e) => Err(e)
                }
                case Err(e) => Err(e)
            }
        )

    pub def seqRight(ma: WbMonad[a], mb: WbMonad[b]): WbMonad[b] = 
        WbMonad(wb -> 
            match apply1(ma, wb) { 
                case Ok(_) => apply1(mb, wb)
                case Err(e) => Err(e)
            }
        )


    pub def bind(ma: WbMonad[a], fn: a -> WbMonad[b]): WbMonad[b] = 
        WbMonad(wb -> 
            match apply1(ma, wb) { 
                case Ok(a) => apply1(fn(a), wb)
                case Err(e) => Err(e)
            }
        )

    pub def flatMap(fn: a -> WbMonad[b], ma: WbMonad[a]): WbMonad[b] = bind(ma, fn)



    def liftWorkbookPrim(fn: ##org.apache.poi.xssf.usermodel.XSSFWorkbook ~> Result[a, System.Error]): WbMonad[a] = 
        WbMonad(wb -> fn(wb))

    pub def getNumberOfSheets(): WbMonad[Int32] =
        use SheetReader/Internal/Prim.{getNumberOfSheets};
        liftWorkbookPrim(x1 -> x1.getNumberOfSheets())

    pub def getSheetName(ix: Int32): WbMonad[String] =
        use SheetReader/Internal/Prim.{getSheetName};
        liftWorkbookPrim(x1 -> x1.getSheetName(ix))        

}