
def main(_args: Array[String]): Int32 & Impure = 
    let src = System/FilePath.new("G:/work/assets/cost-centers/FLOC Upload 26 03.xlsx");
    Console.printLine("Start ${Time/LocalTime.now()}...");
    let _ = Result.getWithDefault(test01(src), 1);
    Result.getWithDefault(test02(src), 1)

def test01(src: System.FilePath): Result[Int32, String] & Impure = 
    use Result.{flatMap};
    let* cr = SheetIO/CursorReader.newRowCursorAt(src, 0);
    let n1  = Cursor.fold((ac,_) -> ac+1, 0, cr);
    let _ = Console.printLine("length: ${n1}");
    Ok(0)


def test02(src: System.FilePath): Result[Int32, String] & Impure = 
    use Result.{flatMap};
    use SheetIO/NamedRowEvaluator.{getContentNamed};
    let* cr = SheetIO/CursorReader.newRowCursorAt(src, 0);
    let cr1 = SheetIO/CursorReader.namedRowCursor(0, getContentNamed("CommonName"), cr);
    let* ans = Cursor.resultFold((ac, s) -> if (String.length(s) > String.length(ac)) s else ac, "", cr1);
    let _ = Console.printLine("longest: ${ans}");
    Ok(0)    