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


    pub type alias RowSetter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.Builder[ka, Interop.Apache.Poi.SS.Usermodel.MutRow[r], ef, a]

}

mod SheetIO.RowSetter {

    use GradedError.{swapError}
    use Graded.Builder.{liftSetter!, liftSetterResult!}

    use SheetIO.RowSetter
    

    use Interop.Apache.Poi.SS.Usermodel.Cell
    use Interop.Apache.Poi.SS.Usermodel.Cell.{Cell}
    use Interop.Apache.Poi.SS.Usermodel.CellStyle
    use Interop.Apache.Poi.SS.Usermodel.CellStyle.{CellStyle}
    use Interop.Apache.Poi.SS.Usermodel.MutRow
    use Interop.Apache.Poi.SS.Usermodel.MutRow.{MutRow}


    /// The state param `st` is expected to be mutated...
    pub def runRowSetter!(ma: RowSetter[a, r, ef, a], 
                            st: MutRow[r]): Result[String, a] \ ef =
        Graded.Builder.runBuilder(ma, st)

/*

    // pub def liftResult1(makeErrMsg: e -> String, x: Result[a, e]): RowSetter[r, a] = match x { 
    //     case Ok(x1) => point(x1)
    //     case Err(e) => throwError(makeErrMsg(e))
    // }

    // pub def liftValidation(makeErrMsg: List[e] -> String, x: Validation[a, e]): RowSetter[r, a] = match x { 
    //     case Success(x1) => point(x1)
    //     case Failure(es) => throwError(makeErrMsg(Nec.toList(es)))
    // }



    /// This also sets the CellStyle...
    def createOrGetCell(ix: Int32): RowSetter[r, ka, Cell[r]] \ r) = 
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
                        f: Option[CellStyle[r]] -> Cell[r] -> a \ ef): RowSetter[r, ka, a] \ { r, ef } =
        let* cell = createOrGetCell(ix);
        liftSetter!((style, _) -> f(style, cell))

    // def liftCellSetterResult!(ix: Int32, 
    //                             f: Option[CellStyle[r]] -> Cell[r] -> Result[a, err] \ ef): RowSetter[r, ka, a] \ { r, ef } with ToString[err] =
    //     let* cell = createOrGetCell(ix);
    //     liftSetterResult!((style, _) -> f(style, cell))


    pub def localCellStyle(style: CellStyle[r], ma: RowSetter[r, ka, a]): RowSetter[r, ka, a] = 
        localEnv(_ -> Some(style), ma)


    pub def putBlankAt!(ix: Int32): RowSetter[r, ka, Unit] \ r =
        liftCellSetter!(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.setBlank!(cell))


    pub def putStringAt!(ix: Int32, value: String): RowSetter[r, ka, Unit] \ r =
        liftCellSetter!(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.setCellValueString!(value, cell))

    pub def putBoolAt!(ix: Int32, value: Bool): RowSetter[r, ka, Unit] \ r =
        liftCellSetter!(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.setCellValueBoolean!(value, cell))


    def putNumericAt!(ix: Int32, value: Float64): RowSetter[r, ka, Unit] \ r =
        liftCellSetter!(ix, (_, cell) -> Interop.Apache.Poi.SS.Usermodel.Cell.setCellValueDouble!(value, cell))

    pub def putInt8At!(ix: Int32, n: Int8): RowSetter[r, ka, Unit] \ r =
        putNumericAt!(ix, Int8.toFloat64(n))

    pub def putInt16At!(ix: Int32, n: Int16): RowSetter[r, ka, Unit] \ r =
        putNumericAt!(ix, Int16.toFloat64(n))

    pub def putInt32At!(ix: Int32, n: Int32): RowSetter[r, ka, Unit] \ r =
        putNumericAt!(ix, Int32.toFloat64(n))


    pub def putInt64At!(ix: Int32, n: Int64): RowSetter[r, ka, Unit] \ r =
        putNumericAt!(ix, Int64.toFloat64(n))

    pub def putFloat32At!(ix: Int32, n: Float32): RowSetter[r, ka, Unit] \ r =
        putNumericAt!(ix, Float32.toFloat64(n))

    pub def putFloat64At!(ix: Int32, n: Float64): RowSetter[r, ka, Unit] \ r =
        putNumericAt!(ix, n)




    // pub def putLocalDateTimeAt!(ix: Int32, dt: Time.LocalDateTime): RowSetter[r, ka, Unit] \ r =
    //     use Time.LocalDateTime.{LocalDateTime};
    //     let  LocalDateTime(dt1) = dt;
    //     liftPrimCell(ix, (r, cell, _) -> SheetIO.Internal.Cell.setCellValueLocalDateTime!(r, dt1, cell))



    // pub def putLocalDateAt!(ix: Int32, dt: Time.LocalDate): RowSetter[r, ka, Unit] \ r =
    //     use Time.LocalDate.{LocalDate};
    //     let  LocalDate(dt1) = dt;
    //     liftPrimCell(ix, (r, cell, _) -> SheetIO.Internal.Cell.setCellValueLocalDate!(r, dt1, cell))
 
 
    pub def putHeaders!(l: List[String]): RowSetter[r, ka, Unit] \ r = 
        putHeadersHelper!(l, 0)
        
    def putHeadersHelper!(l: List[String], ix: Int32): RowSetter[r, ka, Unit] \ r = match l {
        case Nil => point()
        case s :: rs => {
            let* _  = putStringAt!(ix, s);
            putHeadersHelper!(rs, ix + 1)
        }
    }

    pub def putOptionStringAt!(ix: Int32, x: Option[String]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(s) => putStringAt!(ix, s)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionBoolAt!(ix: Int32, x: Option[Bool]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(b) => putBoolAt!(ix, b)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt8At!(ix: Int32, x: Option[Int8]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt8At!(ix, n)
        case None    => putBlankAt!(ix)
        }

    pub def putOptionInt16At!(ix: Int32, x: Option[Int16]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt16At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt32At!(ix: Int32, x: Option[Int32]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt32At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionInt64At!(ix: Int32, x: Option[Int64]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(n) => putInt64At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionFloat32At!(ix: Int32, x: Option[Float32]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(n) => putFloat32At!(ix, n)
        case None    => putBlankAt!(ix)
    }

    pub def putOptionFloat64At!(ix: Int32, x: Option[Float64]): RowSetter[r, ka, Unit] \ r = match x {
        case Some(n) => putFloat64At!(ix, n)
        case None    => putBlankAt!(ix)
    }
*/
}

