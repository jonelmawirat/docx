package docx

import (
    "archive/zip"
    "bytes"
    "encoding/xml"
    "fmt"
    "io"
    "os"

)


type relationship struct {
    XMLName    xml.Name `xml:"Relationship"`
    ID         string   `xml:"Id,attr"`
    Type       string   `xml:"Type,attr"`
    Target     string   `xml:"Target,attr"`
    TargetMode string   `xml:"TargetMode,attr,omitempty"`
}


type relationships struct {
    XMLName       xml.Name       `xml:"Relationships"`
    Xmlns         string         `xml:"xmlns,attr"`
    Relationships []relationship `xml:"Relationship"`
}


type defaultType struct {
    XMLName     xml.Name `xml:"Default"`
    Extension   string   `xml:"Extension,attr"`
    ContentType string   `xml:"ContentType,attr"`
}


type overrideType struct {
    XMLName     xml.Name `xml:"Override"`
    PartName    string   `xml:"PartName,attr"`
    ContentType string   `xml:"ContentType,attr"`
}


type types struct {
    XMLName   xml.Name       `xml:"Types"`
    Xmlns     string         `xml:"xmlns,attr"`
    Defaults  []defaultType  `xml:"Default"`
    Overrides []overrideType `xml:"Override"`
}


type DocumentWriter interface {
    WriteDocument(filename string, doc Document) error
}


type ZipDocxWriter struct{}


func NewZipDocxWriter() *ZipDocxWriter {
    return &ZipDocxWriter{}
}


func (zw *ZipDocxWriter) addXMLPart(zipWriter *zip.Writer, filename string, data interface{}) error {
    partWriter, err := zipWriter.Create(filename)
    if err != nil {
        return fmt.Errorf("failed to create %s in zip: %w", filename, err)
    }

    _, err = partWriter.Write([]byte(xml.Header))
    if err != nil {
        return fmt.Errorf("failed to write xml header for %s: %w", filename, err)
    }

    encoder := xml.NewEncoder(partWriter)
    encoder.Indent("", "  ")
    err = encoder.Encode(data)
    if err != nil {
        return fmt.Errorf("failed to encode xml for %s: %w", filename, err)
    }
    return nil
}


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


func (zw *ZipDocxWriter) writeBytesPart(zipWriter *zip.Writer, filename string, content []byte) error {
    fmt.Printf("DEBUG: writeBytesPart called for: '%s'\n", filename)
    partWriter, err := zipWriter.Create(filename)
    if err != nil {
         fmt.Fprintf(os.Stderr, "ERROR in writeBytesPart (Create) for %s: %v\n", filename, err)
        return fmt.Errorf("failed to create %s in zip: %w", filename, err)
    }
    n, err := io.Copy(partWriter, bytes.NewReader(content))
    if err != nil {
         fmt.Fprintf(os.Stderr, "ERROR in writeBytesPart (Copy) for %s after writing %d bytes: %v\n", filename, n, err)
        return fmt.Errorf("failed to write byte content to %s: %w", filename, err)
    }
     fmt.Printf("DEBUG: writeBytesPart successfully copied %d bytes to '%s'\n", n, filename)
    return nil
}


func (zw *ZipDocxWriter) WriteDocument(filename string, doc Document) error {

    file, err := os.Create(filename)
    if err != nil {
        return fmt.Errorf("failed to create file %s: %w", filename, err)
    }
    defer file.Close()


    zipWriter := zip.NewWriter(file)
    defer zipWriter.Close()


    contentTypes := types{
        Xmlns: "http:
        Defaults: []defaultType{

            {Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
            {Extension: "xml", ContentType: "application/xml"},
        },
        Overrides: []overrideType{

            {PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
            {PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},

        },
    }



    addedExtensions := make(map[string]bool)
    for _, d := range contentTypes.Defaults {
        addedExtensions[d.Extension] = true
    }
    for contentType, ext := range doc.getImageContentTypes() {
        if !addedExtensions[ext] {
            contentTypes.Defaults = append(contentTypes.Defaults, defaultType{Extension: ext, ContentType: contentType})
            addedExtensions[ext] = true
        }
    }


    err = zw.addXMLPart(zipWriter, "[Content_Types].xml", contentTypes)
    if err != nil {
        return fmt.Errorf("failed writing [Content_Types].xml: %w", err)
    }


    rootRels := relationships{
        Xmlns: "http:
        Relationships: []relationship{
            {
                ID:     "rId1",
                Type:   "http:
                Target: "word/document.xml",
            },

        },
    }
    err = zw.addXMLPart(zipWriter, "_rels/.rels", rootRels)
    if err != nil {
        return fmt.Errorf("failed writing _rels/.rels: %w", err)
    }



    docRelsList := []relationship{
        {
            ID:     "rId1",
            Type:   "http:
            Target: "styles.xml",
        },

    }

    docRelsList = append(docRelsList, doc.getImageRelationships()...)

    docRels := relationships{
        Xmlns:         "http:
        Relationships: docRelsList,
    }
    err = zw.addXMLPart(zipWriter, "word/_rels/document.xml.rels", docRels)
    if err != nil {
        return fmt.Errorf("failed writing word/_rels/document.xml.rels: %w", err)
    }


    err = zw.writeStringPart(zipWriter, "word/styles.xml", defaultStylesXML)
    if err != nil {
        return fmt.Errorf("failed writing word/styles.xml: %w", err)
    }


    for imgFilename, imgBytes := range doc.getImages() {

        mediaPath := "word/media/" + imgFilename
        err = zw.writeBytesPart(zipWriter, mediaPath, imgBytes)
        if err != nil {

            return fmt.Errorf("failed to write image %s to zip path %s: %w", imgFilename, mediaPath, err)
        }
    }


    docPartWriter, err := zipWriter.Create("word/document.xml")
    if err != nil {
        return fmt.Errorf("failed to create word/document.xml in zip: %w", err)
    }
    err = doc.renderContent(docPartWriter)
    if err != nil {
        return fmt.Errorf("failed to render document content to zip: %w", err)
    }


    return nil
}
