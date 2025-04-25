package main

import (
    "fmt"
    "os"
    // Assuming your package is in a subdirectory named 'docx'
    // relative to your project's module root, or adjust the import path
    // e.g., "your_module_name/docx"
    "github.com/jonelmawirat/docx" // Use the correct import path for your setup
)

func main() {
    // Create a new document object
    doc := docx.NewDocxDocument()
    // Create a writer object
    writer := docx.NewZipDocxWriter()

    outputFilename := "generated_from_package.docx"

    // Add content
    doc.AddText(docx.StyleHeading1, "Document Created Using Package", docx.FormatItalic)
    // doc.AddNewLine() // AddNewLine adds an empty paragraph. Often not needed after a heading.

    doc.AddText(docx.StyleHeading2, "Introduction", docx.FormatBold)
    // doc.AddNewLine()

    doc.AddText(docx.StyleNormal, "This document demonstrates the direct usage of the docx package.")
    // Test appending runs to the same paragraph
    doc.AddText(docx.StyleNormal, " Multiple calls to AddText with the same style", docx.FormatBold)
    doc.AddText(docx.StyleNormal, " append runs to the same paragraph if possible.")
    doc.AddNewLine() // Add a line break before the image section

    // Define image path (relative to where you run the program)
    // Make sure 'sample/test.png' exists or change the path
    imagePath := "sample/test.png" // Adjust if your image is elsewhere

    // Check if the image file exists
    if _, err := os.Stat(imagePath); os.IsNotExist(err) {
        fmt.Printf("Warning: Test image '%s' not found. Skipping image addition.\n", imagePath)
    } else {
        fmt.Printf("Attempting to add image: %s\n", imagePath)
        err := doc.AddImage(imagePath) // Add the image
        if err != nil {
            // Print a more detailed error if image addition fails
            fmt.Fprintf(os.Stderr, "Error adding image '%s': %v\n", imagePath, err)
        } else {
            fmt.Printf("Successfully added image: %s\n", imagePath)
            // Add text *after* the image (this will be in a new paragraph)
            doc.AddText(docx.StyleNormal, "An image should be displayed above this text.", docx.FormatItalic)
            doc.AddNewLine()
        }
    }

    doc.AddText(docx.StyleHeading2, "Conclusion", docx.FormatBold)
    // doc.AddNewLine()

    doc.AddText(docx.StyleNormal, "The package structure allows for generating DOCX files.")

    // Write the document to the specified file
    fmt.Printf("Writing document to %s...\n", outputFilename)
    err := writer.WriteDocument(outputFilename, doc)
    if err != nil {
        fmt.Fprintf(os.Stderr, "Error writing document '%s': %v\n", outputFilename, err)
        os.Exit(1) // Exit with an error code
    }

    fmt.Printf("Docx file successfully written to %s\n", outputFilename)
}
