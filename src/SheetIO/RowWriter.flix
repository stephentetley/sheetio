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


/// TODO - Use `Option[CellStyle]` as a Reader env in the Monad to allow inherited styles?

mod SheetIO {


    pub type alias RowWriter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.Builder[ka, Interop.Apache.Poi.SS.Usermodel.MutRow[r], ef, a]

}

mod SheetIO.RowWriter {

    use GradedError.{swapError}
    use Graded.Builder.{liftSetter!, liftSetterResult!}

    use SheetIO.RowWriter
    

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.Cell.{Cell}
    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Interop.Apache.Poi.SS.Usermodel.MutRow
    use Interop.Apache.Poi.SS.Usermodel.MutRow.{MutRow}


    /// The state param `st` is expected to be mutated...
    pub def runRowWriter!(ma: RowWriter[a, r, ef, a], 
                            st: MutRow[r]): Result[String, a] \ ef =
        Graded.Builder.runBuilder(ma, st)

/*

    // pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): RowWriter[r, a] = match x { 
    //     case Ok(x1) => point(x1)
    //     case Err(e) => throwError(makeErrMsg(e))
    // }

    // pub def liftValidation(makeErrMsg: List[e] -> String, x: Validation[a, e]): RowWriter[r, a] = match x { 
    //     case Success(x1) => point(x1)
    //     case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    // }



    /// This also sets the CellStyle...
    def createOrGetCell(ix: Int32): RowWriter[r, ka, Cell[r]] \ r) = 
        use Result.{flatMap};
        use Interop.Apache.Poi.SS.Usermodel.Row.{createCell!};
        let body = (style, row) -> {
            let* cell   = match Interop.Apache.Poi.SS.Usermodel.Row.getCell(ix, row) {
                            case Some(ans)  => if (not Object.isNull(ans)) Ok(ans) else createCell!(ix, row)
                            case None       => createCell!(ix, row)
                        };
            let _       = match style { 
                            case Some(sty1) => Interop.Apache.Poi.SS.Usermodel.Cell.setCellStyle!(sty1, cell)
                            case None       => unsafe_cast () as _ \ r
                        };
            Ok(cell)
        };
        liftSetterResult!(body)

    def liftCellSetter!(ix: Int32, 
                        f: Option[CellStyle[r]] -> Cell[r] -> a \ ef): RowWriter[r, ka, a] \ { r, ef } =
        let* cell = createOrGetCell(ix);
        liftSetter!((style, _) -> f(style, cell))

    // def liftCellSetterResult!(ix: Int32, 
    //                             f: Option[CellStyle[r]] -> Cell[r] -> Result[a, err] \ ef): RowWriter[r, ka, a] \ { r, ef } with ToString[err] =
    //     let* cell = createOrGetCell(ix);
    //     liftSetterResult!((style, _) -> f(style, cell))


    pub def localCellStyle(style: CellStyle[r], ma: RowWriter[r, ka, a]): RowWriter[r, ka, a] = 
        localEnv(_ -> Some(style), ma)




 
    pub def putHeaders!(l: List[String]): RowWriter[r, ka, Unit] \ r = 
        putHeadersHelper!(l, 0)
        
    def putHeadersHelper!(l: List[String], ix: Int32): RowWriter[r, ka, Unit] \ r = match l {
        case Nil => point()
        case s :: rs => {
            let* _  = putStringAt!(ix, s);
            putHeadersHelper!(rs, ix + 1)
        }
    }

    pub def putOptionStringAt!(ix: Int32, x: Option[String]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(s) => putStringAt!(ix, s)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionBoolAt!(ix: Int32, x: Option[Bool]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(b) => putBoolAt!(ix, b)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt8At!(ix: Int32, x: Option[Int8]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt8At!(ix, n)
        case None    => putBlankAt!(ix)
        }

    pub def putOptionInt16At!(ix: Int32, x: Option[Int16]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt16At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt32At!(ix: Int32, x: Option[Int32]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt32At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt64At!(ix: Int32, x: Option[Int64]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt64At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionFloat32At!(ix: Int32, x: Option[Float32]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(n) => putFloat32At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionFloat64At!(ix: Int32, x: Option[Float64]): RowWriter[r, ka, Unit] \ r = match x {
        case Some(n) => putFloat64At!(ix, n)
        case None    => putBlankAt!(ix)
    }
*/
}

