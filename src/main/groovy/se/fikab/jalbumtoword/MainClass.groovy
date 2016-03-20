package se.fikab.jalbumtoword

//@Grab(group='com.craigburke.document', module='pdf', version='0.1.6')
//@Grab(group='com.craigburke.document', module='word', version='0.1.6')

import com.craigburke.document.builder.PdfDocumentBuilder
import com.craigburke.document.builder.WordDocumentBuilder

class MainClass {

    public static void main(String... args) {
        System.out.println("Hello my world!")

        new ExcelBuilder("/Users/fredrik/callista/dev/jalbumtoword/testfiles/fillista.xls").eachLine([labels:true]) {
//            new Person(name:"$firstname $lastname",
//                    address:address, telephone:phone).save()
            println "First column on row ${it.rowNum} = ${cell(0)}"
        }


        def builders = [
                new PdfDocumentBuilder(new File('example.pdf')),
                new WordDocumentBuilder(new File('example.docx')),
        ]

        def RAINBOW_COLORS = ['#FF0000', '#FF7F00', '#FFFF00', '#00FF00', '#0000FF', '#4B0082', '#8B00FF']

        String GROOVY_IMAGE_URL = 'http://www.craigburke.com/images/posts/groovy-logo.png'
        byte[] groovyImageData = new URL(GROOVY_IMAGE_URL).bytes

        builders.each { builder ->
            builder.create { document(font: [family: 'Helvetica', size: 14.pt], margin: [top: 0.75.inches]) {
                paragraph "Groovy Document Builder", font: [size: 22.pt]

                paragraph {
                    font.size = 42.pt
                    "Hello Woooorld!!!!!".toUpperCase().eachWithIndex { letter, index ->
                        font.size--
                        font.color = RAINBOW_COLORS[ index % RAINBOW_COLORS.size() ]
                        text letter
                    }
                    lineBreak()
                    text "Current font size is ${font.size}pt"
                }

                paragraph "Font size is back to 14pt now with the default black font"

                paragraph(margin: [left: 1.25.inches, right: 1.inch, top: 0.25.inches, bottom: 0.25.inches]) {
                    font << [family: 'Times-Roman', bold: true, italic: true, color: '#333333']
                    text "A paragraph with a different font and margins"
                }

                paragraph(margin: [left: 1.inch]) {
                    image(data: groovyImageData, width: 250.px, height: 125.px)
                    lineBreak()
                    text "Figure 1: Groovy Logo", font: [italic: true, size: 9.pt]
                }

                paragraph("Suddenly, a table...", font: [size: 22.pt], margin: [bottom: 0.25.inches])

                table(width: 5.inches) {
                    row {
                        cell("Cell 1", width: 1.inch)
                        cell("Cell 2", width: 2.inches)
                        cell(width: 2.inches) {
                            text "Cell 3"
                        }
                    }
                }
            }}
        }


    }

}
