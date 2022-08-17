# sheetio

A Flix library built on top of Apache POI to read and write Excel files.

`StreamingReader` is an experimental streaming reader for large files with homogeneous rows (it is currently broken).

## Flix dependencies 

`pkg` files included in the folder `lib`:

https://github.com/stephentetley/flix-time

https://github.com/stephentetley/basis-base

https://github.com/stephentetley/withindex-classes

https://github.com/stephentetley/monad-lib

https://github.com/stephentetley/interop-base (plus jar)


## Java dependencies 

sheetio needs these jars (mostly from org.apache):

* commons-io
* commons-codec
* commons-collections
* commons-compress
* poi
* poi-ooxml
* poi-ooxml-full
* SparseBitSet
* xmlbeans

Not currently used `excel-streaming-reader-3.1.1.jar` [com.github.pjfanning]

SheetIO includes some Java wrappers that are compiled into `sheetiojava-1.n.jar`.

`sheetiojava-1.n.jar` is a jar of the Java sources in `src/java`.

