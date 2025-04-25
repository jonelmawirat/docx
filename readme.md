# docx

A Go package for programmatically creating Microsoft Word (.docx) documents with support for styled text and image insertion.

## Overview

The `docx` package allows developers to generate DOCX files in Go. It supports adding paragraphs with predefined styles (e.g., Normal, Heading1, Heading2), applying text formatting (bold, italic), and embedding images (JPEG, PNG, GIF). The package creates a valid DOCX file structure, including XML documents, relationships, content types, and media files, using the Open XML format.

Key features:
- Add text with styles (`Normal`, `Heading1`, `Heading2`, `Heading3`, `Heading4`)
- Apply text formatting (bold, italic)
- Insert images with automatic sizing
- Generate valid DOCX files with minimal dependencies

## Installation

To use this package, ensure you have Go 1.24.2 or later installed. Add the package to your project using:

```bash
go get github.com/jonelmawirat/docx
```

The package has no external dependencies beyond the Go standard library.

## Usage

Below is an example of how to use the `docx` package to create a DOCX file with styled text and an image.

### Example

```go
package main

import (
    "fmt"
    "os"
    "github.com/jonelmawirat/docx"
)

func main() {
    // Create a new document
    doc := docx.NewDocxDocument()
    writer := docx.NewZipDocxWriter()

    outputFilename := "example.docx"

    // Add styled text
    doc.AddText(docx.StyleHeading1, "Sample Document", docx.FormatItalic)
    doc.AddText(docx.StyleHeading2, "Introduction", docx.FormatBold)
    doc.AddText(docx.StyleNormal, "This is a sample document created using the docx package.")
    doc.AddText(docx.StyleNormal, " It supports multiple styles and formatting.", docx.FormatBold)
    doc.AddNewLine()

    // Add an image
    imagePath := "sample/test.png"
    if _, err := os.Stat(imagePath); os.IsNotExist(err) {
        fmt.Printf("Warning: Image '%s' not found.\n", imagePath)
    } else {
        err := doc.AddImage(imagePath)
        if err != nil {
            fmt.Fprintf(os.Stderr, "Error adding image: %v\n", err)
        } else {
            doc.AddText(docx.StyleNormal, "Image inserted above.", docx.FormatItalic)
            doc.AddNewLine()
        }
    }

    // Add more text
    doc.AddText(docx.StyleHeading2, "Conclusion", docx.FormatBold)
    doc.AddText(docx.StyleNormal, "The docx package simplifies DOCX file generation.")

    // Write the document
    err := writer.WriteDocument(outputFilename, doc)
    if err != nil {
        fmt.Fprintf(os.Stderr, "Error writing document: %v\n", err)
        os.Exit(1)
    }

    fmt.Printf("Document successfully written to %s\n", outputFilename)
}
```

### Steps to Run

1. Ensure the image file (e.g., `sample/test.png`) exists in the specified path.
2. Run the program:
   ```bash
   go run main.go
   ```
3. Open the generated `example.docx` file in Microsoft Word or a compatible viewer.

### Supported Styles and Formats

- **Paragraph Styles**:
  - `StyleNormal`: Default paragraph style
  - `StyleHeading1`: Level 1 heading
  - `StyleHeading2`: Level 2 heading
  - `StyleHeading3`: Level 3 heading
  - `StyleHeading4`: Level 4 heading

- **Text Formats**:
  - `FormatBold`: Bold text
  - `FormatItalic`: Italic text

### Image Support

- Supported formats: JPEG, PNG, GIF
- Images are embedded in the `word/media/` directory of the DOCX file
- Image dimensions are automatically calculated and converted to EMUs (English Metric Units) for proper scaling

## Project Structure

The package is organized into the following files:

- `document.go`: Defines the `DocxDocument` struct and methods for adding text, images, and rendering content.
- `image_structs.go`: Contains XML structs for image embedding in DOCX files.
- `styles.go`: Provides default Word styles (e.g., Normal, Heading1) as XML.
- `writer.go`: Implements the `ZipDocxWriter` for creating the DOCX ZIP archive.

## Requirements

- Go 1.24.2 or later
- No external dependencies

## Limitations

- Currently supports basic text styling and image insertion. Advanced features like tables, lists, or custom styles are not implemented.
- Only a subset of Word styles is predefined (`Normal`, `Heading1`–`Heading4`).
- Image support is limited to JPEG, PNG, and GIF formats.

