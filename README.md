Java XLS Image Extractor Sample
===============================

Extracting optional images from an .xls file using Apache POI

## Run it:

    ./gradlew clean run

## Output:

    $ ./gradlew clean run
    :clean
    :compileJava
    :processResources
    :classes
    :run
    Total elements: 7
    Row: XlsRow{rowNo='001', title='54 Strat', filename='1954.jpg', imageData=[B@5400d826}
    Row: XlsRow{rowNo='002', title='69 Strat', filename='null', imageData=null}
    Row: XlsRow{rowNo='003', title='79 Strat', filename='1979.jpg', imageData=[B@36ed1e0}
    Row: XlsRow{rowNo='004', title='94 Strat', filename='1994.jpg', imageData=[B@6094cae2}
    Row: XlsRow{rowNo='005', title='Candy Apple Red', filename='CAR.png', imageData=[B@4893ecf7}
    Row: XlsRow{rowNo='006', title='80's MIJ', filename='null', imageData=null}
    Row: XlsRow{rowNo='007', title='John Mayer Sig', filename='JM.png', imageData=[B@67aa715a}

    BUILD SUCCESSFUL

    Total time: 5.865 secs
