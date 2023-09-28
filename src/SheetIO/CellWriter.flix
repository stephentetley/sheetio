/*
 * Copyright 2023 Stephen Tetley
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



mod SheetIO {

    pub type alias CellWriter[ka: Type, r: Region, ef: Eff, a: Type] = 
        Graded.MutStateError[ka, Internal.Apache.Poi.SS.Usermodel.MutCell[r], ef, a]
}

mod SheetIO.CellWriter {

    use Graded.MutStateError.{liftSetter!, liftSetterResult!}

    use SheetIO.CellWriter
    use Internal.Apache.Poi.SS.Usermodel.MutCell
    
    pub def runCellWriter(ma: CellWriter[a, r, ef, a], 
                            st: MutCell[r]): Result[String, a] \ ef =
        Graded.MutStateError.runMutStateError(ma, st)


    pub def putBlank!(): CellWriter[ka, r, ef + r, Unit] =
        liftSetter!(cell -> checked_ecast(Internal.Apache.Poi.SS.Usermodel.MutCell.setBlank!(cell)))


    pub def putString!(value: String): CellWriter[ka, r, ef + r, Unit] =
        liftSetter!(cell -> checked_ecast(Internal.Apache.Poi.SS.Usermodel.MutCell.setCellValueString!(value, cell)))

    pub def putBool!(value: Bool): CellWriter[ka, r, ef + r, Unit] =
        liftSetter!(cell -> checked_ecast(Internal.Apache.Poi.SS.Usermodel.MutCell.setCellValueBoolean!(value, cell)))


    def putNumeric!(value: Float64): CellWriter[ka, r, ef + r, Unit] =
        liftSetter!(cell -> checked_ecast(Internal.Apache.Poi.SS.Usermodel.MutCell.setCellValueDouble!(value, cell)))

    pub def putInt8!(n: Int8): CellWriter[ka, r, ef + r, Unit] =
        putNumeric!(Int8.toFloat64(n))

    pub def putInt16!(n: Int16): CellWriter[ka, r, ef + r, Unit] =
        putNumeric!(Int16.toFloat64(n))

    pub def putInt32!(n: Int32): CellWriter[ka, r, ef + r, Unit] =
        putNumeric!(Int32.toFloat64(n))

    pub def putInt64!(n: Int64): CellWriter[ka, r, ef + r, Unit] =
        putNumeric!(Int64.toFloat64(n))

    pub def putFloat32!(n: Float32): CellWriter[ka, r, ef + r, Unit] =
        putNumeric!(Float32.toFloat64(n))

    pub def putFloat64!(n: Float64): CellWriter[ka, r, ef + r, Unit] =
        putNumeric!(n)

    pub def putLocalDateTime!(dt: Time.LocalDateTime): CellWriter[ka, r, ef + r, Unit] =
        use Time.LocalDateTime.{LocalDateTime};
        let  LocalDateTime(dt1) = dt;
        liftSetter!(cell -> checked_ecast(Internal.Apache.Poi.SS.Usermodel.MutCell.setCellValueLocalDateTime!(dt1, cell)))



    pub def putLocalDate!(dt: Time.LocalDate): CellWriter[ka, r, ef + r, Unit] =
        use Time.LocalDate.{LocalDate};
        let  LocalDate(dt1) = dt;
        liftSetter!(cell -> checked_ecast(Internal.Apache.Poi.SS.Usermodel.MutCell.setCellValueLocalDate!(dt1, cell)))
 

    pub def putOptionString!(x: Option[String]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(s) => putString!(s)
        case None    => putBlank!()
    }

    pub def putOptionBool!(x: Option[Bool]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(b) => putBool!(b)
        case None    => putBlank!()
    }

    pub def putOptionInt8!(x: Option[Int8]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(n) => putInt8!(n)
        case None    => putBlank!()
        }

    pub def putOptionInt16!(x: Option[Int16]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(n) => putInt16!(n)
        case None    => putBlank!()
    }

    pub def putOptionInt32!(x: Option[Int32]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(n) => putInt32!(n)
        case None    => putBlank!()
    }

    pub def putOptionInt64!(x: Option[Int64]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(n) => putInt64!(n)
        case None    => putBlank!()
    }

    pub def putOptionFloat32!(x: Option[Float32]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(n) => putFloat32!(n)
        case None    => putBlank!()
    }

    pub def putOptionFloat64!(x: Option[Float64]): CellWriter[ka, r, ef + r, Unit] = match x {
        case Some(n) => putFloat64!(n)
        case None    => putBlank!()
    }

}
