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
            datas.add(new ImageData(path:Fil, name:Tempnamn, number:it.rowNum, comment:Kommentar))
            println "Läst rad. tnamn: $Tempnamn , fil: $Fil , kommentar: $Kommentar"
        }

        def builder = generateWordDoc ? new WordDocumentBuilder(new File('example.docx')) : new PdfDocumentBuilder(new File('example.pdf'))
        createDocument(builder, datas)

    }

    private static void createDocument(DocumentBuilder builder, List<ImageData> datas) {

        def isPdfBuilder = builder instanceof PdfDocumentBuilder;

        builder.create {
            document(font: [family: 'Helvetica', size: 14.pt], margin: [top: 0.75.inches]) {

                datas.each {
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
                }

            }
        }
    }

    private static String getTextComment(Path path) {
        def commentsFilePath = Paths.get(path.parent.toString() + File.separator + "comments.txt")
        def file = commentsFilePath.toFile()
        if (!file.exists()) {
            return "Ingen JAlbum kommentarsfil funnen för denna bild..."
        }
        def foundRow = file.find { String line ->
            line.startsWith(path.fileName.toString())
        }
        if (foundRow != null) {
            return foundRow.replace(path.fileName.toString() + "=", "")
        }
        return "Ingen kommentar från JAlbum funnen för denna bild..."
    }

}
