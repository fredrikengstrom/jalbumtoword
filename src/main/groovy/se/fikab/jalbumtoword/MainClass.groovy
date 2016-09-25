package se.fikab.jalbumtoword

import com.craigburke.document.builder.PdfDocumentBuilder
import com.craigburke.document.builder.WordDocumentBuilder
import com.craigburke.document.core.builder.DocumentBuilder

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

class MainClass {

    public static void main(String... args) {
        def inputfileArg = args.find {
            it.matches(".*.xls")
        }
//        def inputfile = inputfileArg != null ? inputfileArg : "fillista.xls"
        def inputfile = inputfileArg != null ? inputfileArg : "/Users/fredrik/callista/dev/jalbumtoword/testfiles/fillista.xls"

        def generateWordDoc = args.any {
            it.matches("-w")
        }

        def showHelp = args.any {
            it.matches("-h")
        }

        if (showHelp) {
            println("Give path to xls-file as argument. Use flag '-w' to generate word documents.")
        }

        def datas = new ArrayList<ImageData>();
        new ExcelBuilder(inputfile).eachLine([labels:true]) {
            println "Läst rad. tnamn: $Tempnamn , fil: $Fil , kommentar: $Kommentar"
            if (Tempnamn == null) {
                println("Saknar namn. Rad ej tillagd.")
            } else if (Fil == null) {
                println("Saknar fil-sökväg. Rad ej tillagd.")
            } else {
                datas.add(new ImageData(path: Fil.replaceAll(/^\"|\"$/, ""), name: Tempnamn, number: it.rowNum, comment: Kommentar))
            }
        }

        datas.each { if (!new File(it.path.toString()).file) {
            println("Fil saknas: " + it.name + ". Kontrollera filnamnet och sökvägen. Avbryter!")
            System.exit(1);
        }}

        def now = new Date()
        def nowString = now.format("yyyyMMdd-HHmmss")

        def combinedBuilder = getBuilder(generateWordDoc, nowString + '-kombinerat')
        createCombinedDocument(combinedBuilder, datas)

        def photoBuilder = getBuilder(generateWordDoc, nowString + '-foton')
        createPhotoDocument(photoBuilder, datas)

        def textBuilder = getBuilder(generateWordDoc, nowString + '-texter')
        createTextDocument(textBuilder, datas)

        copyImages(datas, nowString)
    }

    static void copyImages(ArrayList<ImageData> imageDatas, String nowString) {
        def folderName = nowString + "Bilder"
        println("Kopierar bilder till katalogen: " + folderName)
        new File(folderName).mkdir()
        for (ImageData imageData : imageDatas) {
            Path sourcePath = Paths.get(imageData.path)
            Files.copy(sourcePath, Paths.get(folderName + File.separator + imageData.name + "." + getFileExtension(sourcePath)))
        }
    }

    private static String getFileExtension(Path path) {
        String name = path.toString();
        try {
            return name.substring(name.lastIndexOf(".") + 1);
        } catch (Exception e) {
            return "jpg";
        }
    }

    private static DocumentBuilder getBuilder(boolean generateWordDoc, String filename) {
        generateWordDoc ? new WordDocumentBuilder(new File(filename + ".docx")) : new PdfDocumentBuilder(new File(filename + ".pdf"))
    }

    private static void createCombinedDocument(DocumentBuilder builder, List<ImageData> datas) {

        def isPdfBuilder = builder instanceof PdfDocumentBuilder;
        println("Skapar kombinerat dokument")
        builder.create {
            document(font: [family: 'Helvetica', size: 14.pt], margin: [top: 0.75.inches]) {

                datas.each {
                    try {
                        Path path = Paths.get(it.path).normalize();
                        byte[] groovyImageData = Files.readAllBytes(path);

                        def textComment = getTextComment(path)
                        def current = it

                        paragraph(margin: [left: 0.inch]) {
                            if (isPdfBuilder) {
                                image(data: groovyImageData, width: 300.px, height: 300.px, name: current.name + current.number)
                            } else {
                                image(data: groovyImageData, name: current.name + current.number)
                            }
                            lineBreak()
                            text "Namn: $current.name", font: [italic: true, size: 9.pt]
                            lineBreak()
                            text "Nummer: $current.number", font: [italic: true, size: 9.pt]
                            lineBreak()
                            text "Sökväg: $current.path", font: [italic: true, size: 9.pt]
                            lineBreak()
                            text "Kommentar: $current.comment", font: [italic: true, size: 9.pt]
                            lineBreak()
                            text "Kommentar från JAlbum: $textComment", font: [italic: true, size: 9.pt]
                            pageBreak()
                        }
                        print(".")
                    } catch (Exception e) {
                        println("Kunde ej lägga till bild: " + it.name)
                    }
                }

            }
        }
    }
        private static void createPhotoDocument(DocumentBuilder builder, List<ImageData> datas) {

        def isPdfBuilder = builder instanceof PdfDocumentBuilder;
            println("Skapar foto-dokument")
        builder.create {
            document(font: [family: 'Helvetica', size: 14.pt], margin: [top: 0.75.inches]) {

                datas.each {
                    Path path = Paths.get(it.path).normalize();
                    byte[] groovyImageData = Files.readAllBytes(path);

//                    def textComment = getTextComment(path)
                    def current = it

                    paragraph(margin: [left: 0.inch]) {
                        if (isPdfBuilder) {
                            image(data: groovyImageData, width: 300.px, height: 300.px, name: current.name + current.number)
                        } else {
                            image(data: groovyImageData, name: current.name + current.number)
                        }
                        lineBreak()
                        text "Namn: $current.name", font: [italic: true, size: 9.pt]
//                        lineBreak()
//                        text "Nummer: $current.number", font: [italic: true, size: 9.pt]
//                        lineBreak()
//                        text "Sökväg: $current.path", font: [italic: true, size: 9.pt]
//                        lineBreak()
//                        text "Kommentar: $current.comment", font: [italic: true, size: 9.pt]
//                        lineBreak()
//                        text "Kommentar från JAlbum: $textComment", font: [italic: true, size: 9.pt]
                        pageBreak()
                    }
                    print(".")
                }

            }
        }
    }

    private static void createTextDocument(DocumentBuilder builder, List<ImageData> datas) {

//        def isPdfBuilder = builder instanceof PdfDocumentBuilder;
        println("Skapar text-dokument")
        builder.create {
            document(font: [family: 'Helvetica', size: 14.pt], margin: [top: 0.75.inches]) {

                datas.each {
                    Path path = Paths.get(it.path).normalize();
//                    byte[] groovyImageData = Files.readAllBytes(path);

                    def textComment = getTextComment(path)
                    def current = it

                    paragraph(margin: [left: 0.inch]) {
//                        if (isPdfBuilder) {
//                            image(data: groovyImageData, width: 300.px, height: 300.px, name: current.name + current.number)
//                        } else {
//                            image(data: groovyImageData, name: current.name + current.number)
//                        }
//                        lineBreak()
                        text "Namn: $current.name", font: [italic: true, size: 9.pt]
                        lineBreak()
                        text "Nummer: $current.number", font: [italic: true, size: 9.pt]
                        lineBreak()
                        text "Sökväg: $current.path", font: [italic: true, size: 9.pt]
                        lineBreak()
                        text "Kommentar: $current.comment", font: [italic: true, size: 9.pt]
                        lineBreak()
                        text "Kommentar från JAlbum: $textComment", font: [italic: true, size: 9.pt]
//                        pageBreak()
                    }
                    print(".")
                }

            }
        }
    }

    private static String getTextComment(Path path) {
        def commentsFilePath = Paths.get(path.parent.toString() + File.separator + "comments.properties")
        def file = commentsFilePath.toFile()
        if (!file.exists()) {
            return "Ingen JAlbum kommentarsfil funnen för denna bild..."
        }
//        final String filetext = file.text;
//        def foundRow = file.find { String line ->
//            line.startsWith(path.fileName.toString())
//        }

//        def lineNo = 1
        boolean includeNextLine = false
        boolean isRowFound = false
        String foundRow
        String line
        def newLine = System.getProperty("line.separator")
        file.withReader { reader ->
            while ((line = reader.readLine())!=null) {
                if (line.startsWith(path.fileName.toString())) {
                    includeNextLine = line.endsWith("\\")
                    foundRow = line.replace("\\", "");
                    isRowFound = true
                } else if (includeNextLine) {
                    includeNextLine = line.endsWith("\\")
                    foundRow += newLine + line.replace("\\", "");
                }
                if (isRowFound && !includeNextLine) {
                    break;
                }
            }
        }

        if (foundRow != null) {
            return foundRow.replace(path.fileName.toString() + "=", "")
        }
        return "Ingen kommentar från JAlbum funnen för denna bild..."
    }

}
