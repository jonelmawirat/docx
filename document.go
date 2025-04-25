package docx

import (
    "bytes" // Import bytes
    "encoding/xml"
    "fmt"
    "image"
    _ "image/gif"
    _ "image/jpeg"
    _ "image/png"
    "io"
    "os"
    "strings"
)

const (
    StyleNormal   = "Normal"
    StyleHeading1 = "Heading1"
    StyleHeading2 = "Heading2"
    StyleHeading3 = "Heading3"
    StyleHeading4 = "Heading4"
    FormatBold    = "Bold"
    FormatItalic  = "Italic"
)

const emusPerPixel = 9525 // Correct constant for typical 96 DPI screen resolution (914400 EMU/inch / 96 pixel/inch)

type boldProperty struct {
    XMLName xml.Name `xml:"w:b"`
}

type italicProperty struct {
    XMLName xml.Name `xml:"w:i"`
}

type runProperties struct {
    XMLName xml.Name        `xml:"w:rPr"`
    Bold    *boldProperty   `xml:"w:b,omitempty"`
    Italic  *italicProperty `xml:"w:i,omitempty"`
    // Add other properties like font, size, color etc. here if needed
}

type paragraphRunText struct {
    XMLName xml.Name `xml:"w:t"`
    Space   string   `xml:"xml:space,attr,omitempty"`
    Text    string   `xml:",chardata"`
}

type paragraphRun struct {
    XMLName    xml.Name          `xml:"w:r"`
    Properties *runProperties    `xml:"w:rPr,omitempty"`
    Break      *struct{ XMLName xml.Name `xml:"w:br"` } `xml:"w:br,omitempty"`
    Drawing    *Drawing          `xml:"w:drawing,omitempty"` // Uses corrected Drawing struct
    Text       *paragraphRunText `xml:"w:t,omitempty"`
}

type paragraphStyle struct {
    XMLName xml.Name `xml:"w:pStyle"`
    Val     string   `xml:"w:val,attr"`
}

type paragraphProperties struct {
    XMLName xml.Name        `xml:"w:pPr"`
    Style   *paragraphStyle `xml:"w:pStyle,omitempty"`
    // Add other paragraph properties like spacing, indentation etc. here if needed
}

type paragraphData struct {
    XMLName    xml.Name             `xml:"w:p"`
    Properties *paragraphProperties `xml:"w:pPr,omitempty"`
    Runs       []paragraphRun       `xml:"w:r,omitempty"`
}

type documentBodyData struct {
    XMLName    xml.Name        `xml:"w:body"`
    Paragraphs []paragraphData `xml:"w:p"`
    SectPr     *sectPr         `xml:"w:sectPr"` // Use pointer for optional section properties
}

// --- Section Properties structs (unchanged, but included for completeness) ---
type pgSz struct {
    XMLName xml.Name `xml:"w:pgSz"`
    W       uint     `xml:"w:w,attr"`
    H       uint     `xml:"w:h,attr"`
}

type pgMar struct {
    XMLName xml.Name `xml:"w:pgMar"`
    Top     int      `xml:"w:top,attr"`
    Right   uint     `xml:"w:right,attr"`
    Bottom  int      `xml:"w:bottom,attr"`
    Left    uint     `xml:"w:left,attr"`
    Header  uint     `xml:"w:header,attr"`
    Footer  uint     `xml:"w:footer,attr"`
    Gutter  uint     `xml:"w:gutter,attr"`
}

type sectPr struct {
    XMLName xml.Name `xml:"w:sectPr"`
    PgSz    pgSz     `xml:"w:pgSz"`
    PgMar   pgMar    `xml:"w:pgMar"`
    Cols    struct {
        XMLName xml.Name `xml:"w:cols"`
        Space   uint     `xml:"w:space,attr"`
    } `xml:"w:cols"`
    DocGrid struct {
        XMLName   xml.Name `xml:"w:docGrid"`
        LinePitch uint     `xml:"w:linePitch,attr"`
    } `xml:"w:docGrid"`
}

// --- Root Document struct (with all required namespaces) ---
type xmlRootDocument struct {
    XMLName xml.Name         `xml:"w:document"`
    XmlnsWp string           `xml:"xmlns:wp,attr"`
    XmlnsA  string           `xml:"xmlns:a,attr"`
    XmlnsPic string          `xml:"xmlns:pic,attr"`
    XmlnsR  string           `xml:"xmlns:r,attr"`
    XmlnsW  string           `xml:"xmlns:w,attr"`
    // Add other namespaces if needed (e.g., xmlns:m, xmlns:v)
    Body documentBodyData `xml:"w:body"`
}

// --- Document Interface (unchanged) ---
type Document interface {
    AddText(style string, textData string, formatOptions ...string)
    AddNewLine()
    AddImage(filepath string) error
    renderContent(w io.Writer) error
    getImages() map[string][]byte
    getImageContentTypes() map[string]string // key=contentType, value=extension
    getImageRelationships() []relationship
}

// --- DocxDocument struct (unchanged) ---
type DocxDocument struct {
    content           []paragraphData
    images            map[string][]byte       // key=filename in word/media/
    imageContentTypes map[string]string       // key=contentType, value=extension (e.g. "image/png": "png")
    imageRels         []relationship          // Relationships for images
    imageCounter      uint                    // Counter for unique image IDs (docPr, cNvPr)
    lastRID           int                     // Counter for unique relationship IDs (rId#)
}

// --- NewDocxDocument constructor ---
func NewDocxDocument() *DocxDocument {
    return &DocxDocument{
        content:           []paragraphData{},
        images:            make(map[string][]byte),
        imageContentTypes: make(map[string]string),
        imageRels:         []relationship{},
        imageCounter:      0, // Start image IDs from 1 if preferred
        lastRID:           1, // Start relationship IDs from rId1 (often used by styles.xml)
    }
}

// --- nextRID helper ---
func (d *DocxDocument) nextRID() string {
    d.lastRID++
    return fmt.Sprintf("rId%d", d.lastRID)
}

// --- AddText (Corrected logic for appending runs) ---
func (d *DocxDocument) AddText(style string, textData string, formatOptions ...string) {
    runProps := runProperties{}
    runText := paragraphRunText{Text: textData, Space: "preserve"} // Use preserve to keep spaces
    var finalRunProps *runProperties
    hasFormatting := false
    for _, opt := range formatOptions {
        if opt == FormatBold {
            runProps.Bold = &boldProperty{}
            hasFormatting = true
        }
        if opt == FormatItalic {
            runProps.Italic = &italicProperty{}
            hasFormatting = true
        }
    }
    if hasFormatting {
        finalRunProps = &runProps
    }

    run := paragraphRun{
        Properties: finalRunProps,
        Text:       &runText,
    }

    // Logic to append to the last paragraph if styles match and it's not an image paragraph
    canAppend := false
    if len(d.content) > 0 {
        lastParaIndex := len(d.content) - 1
        lastPara := &d.content[lastParaIndex]

        // Check if last paragraph is suitable for appending
        // It must have runs OR properties (to avoid appending to empty <w:p/> from AddNewLine)
        // It must not contain a drawing in its runs
        hasDrawing := false
        for _, r := range lastPara.Runs {
            if r.Drawing != nil {
                hasDrawing = true
                break
            }
        }

        if !hasDrawing && (len(lastPara.Runs) > 0 || lastPara.Properties != nil) {
            lastStyle := StyleNormal // Default if no properties
            if lastPara.Properties != nil && lastPara.Properties.Style != nil {
                lastStyle = lastPara.Properties.Style.Val
            }

            currentStyle := style
            if currentStyle == "" { // Treat empty style as Normal for comparison
                currentStyle = StyleNormal
            }

            if currentStyle == lastStyle {
                canAppend = true
            }
        }
    }

    if canAppend {
        lastPara := &d.content[len(d.content)-1]
        lastPara.Runs = append(lastPara.Runs, run)
    } else {
        // Create a new paragraph
        paraProps := paragraphProperties{}
        var finalParaProps *paragraphProperties

        // Determine the style for the new paragraph
        validStyle := style
        styleIsSet := false
        switch style {
        case StyleHeading1, StyleHeading2, StyleHeading3, StyleHeading4:
            paraProps.Style = &paragraphStyle{Val: style}
            finalParaProps = &paraProps
            styleIsSet = true
        case StyleNormal, "":
            validStyle = StyleNormal // Ensure Normal is used if "" provided
            // Don't explicitly set Normal style unless necessary (it's default)
            // But if a specific style was requested (even "Normal"), set it.
            if style == StyleNormal {
                paraProps.Style = &paragraphStyle{Val: StyleNormal}
                finalParaProps = &paraProps
                styleIsSet = true
            }
        default:
            // Handle potentially custom styles
            if strings.TrimSpace(style) != "" {
                paraProps.Style = &paragraphStyle{Val: style}
                finalParaProps = &paraProps
                styleIsSet = true
            }
        }

        // If no valid style was set and default isn't Normal, explicitly set Normal
        if !styleIsSet && validStyle != StyleNormal {
            paraProps.Style = &paragraphStyle{Val: StyleNormal}
            finalParaProps = &paraProps
        }

        para := paragraphData{
            Properties: finalParaProps,
            Runs:       []paragraphRun{run},
        }
        d.content = append(d.content, para)
    }
}

// --- AddNewLine ---
func (d *DocxDocument) AddNewLine() {
    // Adds an empty paragraph, which Word interprets as a line break.
    // Ensure AddText logic doesn't append to this empty paragraph accidentally.
    d.content = append(d.content, paragraphData{})
}

// --- AddImage (Corrected) ---
func (d *DocxDocument) AddImage(filePath string) error {
    d.imageCounter++
    imgID := d.imageCounter // Use this for drawing element IDs
    uniquePicID := imgID     // Use the same counter for non-visual props ID for simplicity
    rID := d.nextRID()       // Generate unique Relationship ID

    imgBytes, err := os.ReadFile(filePath)
    if err != nil {
        return fmt.Errorf("failed to read image file %s: %w", filePath, err)
    }

    // Use bytes.Reader for decoding functions
    imgDataReader := bytes.NewReader(imgBytes)

    imgConfig, format, err := image.DecodeConfig(imgDataReader)
    if err != nil {
        // Reset reader and try full decode if config fails (e.g., some PNGs)
        _, _ = imgDataReader.Seek(0, io.SeekStart) // Reset reader position
        _, format, err = image.Decode(imgDataReader)
        if err != nil {
            return fmt.Errorf("failed to decode image config or data for %s: %w", filePath, err)
        }
        // If Decode succeeded, reset reader again and get config
        _, _ = imgDataReader.Seek(0, io.SeekStart)
        imgConfig, _, err = image.DecodeConfig(imgDataReader)
        if err != nil {
            return fmt.Errorf("failed to get image config after successful decode for %s: %w", filePath, err)
        }
    }

    contentType := ""
    imgExt := ""
    switch format {
    case "jpeg":
        contentType = "image/jpeg"
        imgExt = ".jpg" // Standard extension
    case "png":
        contentType = "image/png"
        imgExt = ".png"
    case "gif":
        contentType = "image/gif"
        imgExt = ".gif"
    default:
        return fmt.Errorf("unsupported image format: %s for file %s", format, filePath)
    }

    // Generate a unique filename for storage within the DOCX
    imgFileName := fmt.Sprintf("image%d%s", imgID, imgExt)

    // Store image data and metadata
    d.images[imgFileName] = imgBytes
    d.imageContentTypes[contentType] = imgExt[1:] // Store mapping: "image/png" -> "png"

    // Add relationship entry
    d.imageRels = append(d.imageRels, relationship{
        ID:     rID,
        Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target: fmt.Sprintf("media/%s", imgFileName), // Use forward slash for Target path
    })

    // Calculate size in EMUs
    widthEMU := int64(imgConfig.Width) * emusPerPixel
    heightEMU := int64(imgConfig.Height) * emusPerPixel

    // Construct the Drawing struct using corrected definitions from image_structs.go
    descr := "Inserted Picture" // Simple description, customize if needed
    drawing := Drawing{
        Inline: Inline{
            DistT: 0, DistB: 0, DistL: 0, DistR: 0,
            Extent:       extent{Cx: widthEMU, Cy: heightEMU},       // <wp:extent>
            EffectExtent: effectExtent{L: 0, T: 0, R: 0, B: 0},      // <wp:effectExtent>
            DocPr: DocProperties{ // <wp:docPr>
                ID:    imgID,
                Name:  fmt.Sprintf("Picture %d", imgID),
                Descr: descr,
            },
            CNvGraphicFramePr: CnvGraphicFrameProperties{ // <wp:cNvGraphicFramePr>
                GraphicFrame: graphicFrameLocks{ // <a:graphicFrameLocks>
                    NoChangeAspect: 1,
                },
            },
            Graphic: graphic{ // <a:graphic>
                GraphicData: graphicData{ // <a:graphicData>
                    URI: "http://schemas.openxmlformats.org/drawingml/2006/picture",
                    Pic: pic{ // <pic:pic>
                        NvPicPr: nonVisualPicProperties{ // <pic:nvPicPr>
                            CNvPr: struct { // <pic:cNvPr>
                                XMLName xml.Name `xml:"pic:cNvPr"`
                                ID      uint     `xml:"id,attr"`
                                Name    string   `xml:"name,attr"`
                                Descr   string   `xml:"descr,attr,omitempty"`
                            }{ID: uniquePicID, Name: imgFileName, Descr: descr}, // Use unique ID
                            CNvPicPr: cNvPicPr{ // <pic:cNvPicPr>
                                PicLocks: picLocks{ // <a:picLocks>
                                    NoChangeAspect:     1,
                                    NoChangeArrowheads: 1, // As per example
                                },
                            },
                        },
                        BlipFill: blipFill{ // <pic:blipFill>
                            Blip: blip{ // <a:blip>
                                Embed:  rID, // Link to relationship ID
                                Cstate: "print",
                            },
                            SrcRect: &srcRect{}, // <a:srcRect/> (empty element)
                            Stretch: stretch{ // <a:stretch>
                                FillRectangle: &struct { // <a:fillRect/> (empty element)
                                    XMLName xml.Name `xml:"a:fillRect"`
                                }{},
                            },
                        },
                        SpPr: shapeProperties{ // <pic:spPr>
                            BwMode: "auto", // As per example
                            Xfrm: transform2D{ // <a:xfrm>
                                Offset:  point2D{X: 0, Y: 0},
                                Extents: extents{Cx: widthEMU, Cy: heightEMU},
                            },
                            PrstGeom: presetGeometry{ // <a:prstGeom>
                                Prst: "rect",
                                AVList: &struct { // <a:avLst/> (empty element)
                                    XMLName xml.Name `xml:"a:avLst"`
                                }{},
                            },
                            NoFill: &noFill{}, // <a:noFill/> (empty element)
                            Ln: &ln{ // <a:ln>
                                NoFill: &noFill{}, // <a:noFill/> within ln (as per example)
                                // Optional: Add Miter, HeadEnd, TailEnd if needed based on full example
                                // Miter: &struct { XMLName xml.Name `xml:"a:miter"`; Lim string `xml:"lim,attr"`}{Lim: "800000"},
                                // HeadEnd: &headEnd{},
                                // TailEnd: &tailEnd{},
                            },
                        },
                    },
                },
            },
        },
    }

    // Create a new paragraph specifically for the image
    imgRun := paragraphRun{Drawing: &drawing}
    // Create a new paragraph containing only the image run
    para := paragraphData{Runs: []paragraphRun{imgRun}}
    d.content = append(d.content, para)

    return nil
}

// --- renderContent ---
func (d *DocxDocument) renderContent(w io.Writer) error {
    // Define standard namespaces
    doc := xmlRootDocument{
        XmlnsWp:  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        XmlnsA:   "http://schemas.openxmlformats.org/drawingml/2006/main",
        XmlnsPic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
        XmlnsR:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        XmlnsW:   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        Body: documentBodyData{
            Paragraphs: d.content,
            // Include default section properties for basic layout
            SectPr: &sectPr{
                PgSz: pgSz{W: 12240, H: 15840}, // Standard Letter size
                PgMar: pgMar{Top: 1440, Right: 1440, Bottom: 1440, Left: 1440, Header: 720, Footer: 720, Gutter: 0}, // 1 inch margins
                Cols: struct {
                    XMLName xml.Name `xml:"w:cols"`
                    Space   uint     `xml:"w:space,attr"`
                }{Space: 720}, // Default column spacing
                DocGrid: struct {
                    XMLName   xml.Name `xml:"w:docGrid"`
                    LinePitch uint     `xml:"w:linePitch,attr"`
                }{LinePitch: 360}, // Default line pitch
            },
        },
    }

    // Write XML header
    _, err := w.Write([]byte(xml.Header))
    if err != nil {
        return fmt.Errorf("failed to write xml header: %w", err)
    }

    // Encode the document structure
    encoder := xml.NewEncoder(w)
    encoder.Indent("", "  ") // Use indentation for readability
    err = encoder.Encode(doc)
    if err != nil {
        return fmt.Errorf("failed to encode document content: %w", err)
    }

    return nil
}

// --- Getter methods (unchanged) ---
func (d *DocxDocument) getImages() map[string][]byte {
    return d.images
}

func (d *DocxDocument) getImageContentTypes() map[string]string {
    return d.imageContentTypes // key=contentType, value=extension
}

func (d *DocxDocument) getImageRelationships() []relationship {
    return d.imageRels
}

