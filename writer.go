package docx

import (
    "archive/zip"
    "bytes"
    "encoding/xml"
    "fmt"
    "io"
    "os"
    // Removed "path" import as we'll use string concat with "/"
)

// --- relationship struct (unchanged) ---
type relationship struct {
    XMLName    xml.Name `xml:"Relationship"`
    ID         string   `xml:"Id,attr"`
    Type       string   `xml:"Type,attr"`
    Target     string   `xml:"Target,attr"`
    TargetMode string   `xml:"TargetMode,attr,omitempty"`
}

// --- relationships struct (unchanged) ---
type relationships struct {
    XMLName       xml.Name       `xml:"Relationships"`
    Xmlns         string         `xml:"xmlns,attr"`
    Relationships []relationship `xml:"Relationship"`
}

// --- defaultType struct (unchanged) ---
type defaultType struct {
    XMLName     xml.Name `xml:"Default"`
    Extension   string   `xml:"Extension,attr"`
    ContentType string   `xml:"ContentType,attr"`
}

// --- overrideType struct (unchanged) ---
type overrideType struct {
    XMLName     xml.Name `xml:"Override"`
    PartName    string   `xml:"PartName,attr"`
    ContentType string   `xml:"ContentType,attr"`
}

// --- types struct (unchanged) ---
type types struct {
    XMLName   xml.Name       `xml:"Types"`
    Xmlns     string         `xml:"xmlns,attr"`
    Defaults  []defaultType  `xml:"Default"`
    Overrides []overrideType `xml:"Override"`
}

// --- DocumentWriter Interface (unchanged) ---
type DocumentWriter interface {
    WriteDocument(filename string, doc Document) error
}

// --- ZipDocxWriter struct (unchanged) ---
type ZipDocxWriter struct{}

// --- NewZipDocxWriter constructor (unchanged) ---
func NewZipDocxWriter() *ZipDocxWriter {
    return &ZipDocxWriter{}
}

// --- addXMLPart helper (unchanged) ---
func (zw *ZipDocxWriter) addXMLPart(zipWriter *zip.Writer, filename string, data interface{}) error {
    partWriter, err := zipWriter.Create(filename)
    if err != nil {
        return fmt.Errorf("failed to create %s in zip: %w", filename, err)
    }
    // Write XML Header
    _, err = partWriter.Write([]byte(xml.Header))
    if err != nil {
        return fmt.Errorf("failed to write xml header for %s: %w", filename, err)
    }
    // Encode XML data
    encoder := xml.NewEncoder(partWriter)
    encoder.Indent("", "  ") // Indent for readability
    err = encoder.Encode(data)
    if err != nil {
        return fmt.Errorf("failed to encode xml for %s: %w", filename, err)
    }
    return nil
}

// --- writeStringPart helper (unchanged) ---
func (zw *ZipDocxWriter) writeStringPart(zipWriter *zip.Writer, filename string, content string) error {
    partWriter, err := zipWriter.Create(filename)
    if err != nil {
        return fmt.Errorf("failed to create %s in zip: %w", filename, err)
    }
    _, err = io.WriteString(partWriter, content)
    if err != nil {
        return fmt.Errorf("failed to write string content to %s: %w", filename, err)
    }
    return nil
}

// --- writeBytesPart helper (unchanged) ---
func (zw *ZipDocxWriter) writeBytesPart(zipWriter *zip.Writer, filename string, content []byte) error {
    partWriter, err := zipWriter.Create(filename)
    if err != nil {
        return fmt.Errorf("failed to create %s in zip: %w", filename, err)
    }
    _, err = io.Copy(partWriter, bytes.NewReader(content))
    if err != nil {
        return fmt.Errorf("failed to write byte content to %s: %w", filename, err)
    }
    return nil
}

// --- WriteDocument (Corrected) ---
func (zw *ZipDocxWriter) WriteDocument(filename string, doc Document) error {
    // Create the output file
    file, err := os.Create(filename)
    if err != nil {
        return fmt.Errorf("failed to create file %s: %w", filename, err)
    }
    defer file.Close() // Ensure file is closed

    // Create a new ZIP archive
    zipWriter := zip.NewWriter(file)
    defer zipWriter.Close() // Ensure zip writer is closed

    // --- 1. [Content_Types].xml ---
    contentTypes := types{
        Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
        Defaults: []defaultType{
            // Basic defaults
            {Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
            {Extension: "xml", ContentType: "application/xml"},
        },
        Overrides: []overrideType{
            // Core document parts
            {PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
            {PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
            // Add other necessary overrides if using features like settings.xml, theme/theme1.xml, webSettings.xml, fontTable.xml etc.
        },
    }

    // Add Default entries for image content types
    // Input map is map[contentType]extension
    addedExtensions := make(map[string]bool) // Track added extensions to avoid duplicates
    for _, d := range contentTypes.Defaults {
        addedExtensions[d.Extension] = true
    }
    for contentType, ext := range doc.getImageContentTypes() {
        if !addedExtensions[ext] {
            contentTypes.Defaults = append(contentTypes.Defaults, defaultType{Extension: ext, ContentType: contentType})
            addedExtensions[ext] = true
        }
    }

    // Write [Content_Types].xml
    err = zw.addXMLPart(zipWriter, "[Content_Types].xml", contentTypes)
    if err != nil {
        return fmt.Errorf("failed writing [Content_Types].xml: %w", err) // Wrap error
    }

    // --- 2. _rels/.rels (Root relationships) ---
    rootRels := relationships{
        Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
        Relationships: []relationship{
            {
                ID:     "rId1", // This ID MUST correspond to the styles.xml relationship ID in docRels if styles.xml is rId1 there. Let's adjust docRels instead.
                Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                Target: "word/document.xml", // Target path uses forward slashes
            },
            // Add relationships for other package-level parts if needed (e.g., core properties)
        },
    }
    err = zw.addXMLPart(zipWriter, "_rels/.rels", rootRels)
    if err != nil {
        return fmt.Errorf("failed writing _rels/.rels: %w", err)
    }

    // --- 3. word/_rels/document.xml.rels (Document relationships) ---
    // Start with the mandatory styles relationship
    docRelsList := []relationship{
        {
            ID:     "rId1", // Start document relationships usually from rId1 for styles
            Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            Target: "styles.xml", // Target relative to word/ directory
        },
        // Add other necessary relationships (fontTable, theme, settings etc.)
    }
    // Append image relationships (which will get IDs rId2, rId3, ...)
    docRelsList = append(docRelsList, doc.getImageRelationships()...)

    docRels := relationships{
        Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
        Relationships: docRelsList,
    }
    err = zw.addXMLPart(zipWriter, "word/_rels/document.xml.rels", docRels) // Path uses forward slashes
    if err != nil {
        return fmt.Errorf("failed writing word/_rels/document.xml.rels: %w", err)
    }

    // --- 4. word/styles.xml ---
    err = zw.writeStringPart(zipWriter, "word/styles.xml", defaultStylesXML) // Path uses forward slashes
    if err != nil {
        return fmt.Errorf("failed writing word/styles.xml: %w", err)
    }

    // --- 5. word/media/ files ---
    for imgFilename, imgBytes := range doc.getImages() {
        // CORRECTED: Use forward slashes ALWAYS for ZIP paths
        mediaPath := "word/media/" + imgFilename
        err = zw.writeBytesPart(zipWriter, mediaPath, imgBytes)
        if err != nil {
            // Provide more context in error message
            return fmt.Errorf("failed to write image %s to zip path %s: %w", imgFilename, mediaPath, err)
        }
    }

    // --- 6. word/document.xml ---
    docPartWriter, err := zipWriter.Create("word/document.xml") // Path uses forward slashes
    if err != nil {
        return fmt.Errorf("failed to create word/document.xml in zip: %w", err)
    }
    err = doc.renderContent(docPartWriter) // Render the main document content
    if err != nil {
        return fmt.Errorf("failed to render document content to zip: %w", err)
    }

    // --- 7. Add other required parts ---
    // A valid DOCX often requires theme, settings, fontTable, etc.
    // For now, we'll skip them, but Word might complain about missing parts or use defaults.

    // zipWriter.Close() is handled by defer
    // file.Close() is handled by defer
    return nil // Success
}
