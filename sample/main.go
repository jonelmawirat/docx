package main

import (
    "fmt"
    "os"
    "github.com/jonelmawirat/docx"
)

func main() {

    doc := docx.NewDocxDocument()
    writer := docx.NewZipDocxWriter()
    outputFilename := "generated_from_package.docx"

    doc.AddText(docx.StyleHeading1, "Document Created Using Package", docx.FormatItalic)
    doc.AddText(docx.StyleHeading2, "Introduction", docx.FormatBold)
    doc.AddText(docx.StyleNormal, "This document demonstrates the direct usage of the docx package.")
    doc.AddText(docx.StyleNormal, " Multiple calls to AddText with the same style", docx.FormatBold)
    doc.AddText(docx.StyleNormal, " append runs to the same paragraph if possible.")
    doc.AddNewLine()



    imagePath := "sample/test.png"


    if _, err := os.Stat(imagePath); os.IsNotExist(err) {
        fmt.Printf("Warning: Test image '%s' not found. Skipping image addition.\n", imagePath)
    } else {
        fmt.Printf("Attempting to add image: %s\n", imagePath)
        err := doc.AddImage(imagePath)
        if err != nil {

            fmt.Fprintf(os.Stderr, "Error adding image '%s': %v\n", imagePath, err)
        } else {
            fmt.Printf("Successfully added image: %s\n", imagePath)

            doc.AddText(docx.StyleNormal, "An image should be displayed above this text.", docx.FormatItalic)
            doc.AddNewLine()
        }
    }

    doc.AddText(docx.StyleHeading2, "Conclusion", docx.FormatBold)
    doc.AddText(docx.StyleNormal, "The package structure allows for generating DOCX files.")


    fmt.Printf("Writing document to %s...\n", outputFilename)
    err := writer.WriteDocument(outputFilename, doc)
    if err != nil {
        fmt.Fprintf(os.Stderr, "Error writing document '%s': %v\n", outputFilename, err)
        os.Exit(1)
    }

    fmt.Printf("Docx file successfully written to %s\n", outputFilename)
}
