package docx

import (
    "bytes"
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

const emusPerPixel = 9525

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
    Drawing    *Drawing          `xml:"w:drawing,omitempty"`
    Text       *paragraphRunText `xml:"w:t,omitempty"`
}

type paragraphStyle struct {
    XMLName xml.Name `xml:"w:pStyle"`
    Val     string   `xml:"w:val,attr"`
}

type paragraphProperties struct {
    XMLName xml.Name        `xml:"w:pPr"`
    Style   *paragraphStyle `xml:"w:pStyle,omitempty"`

}

type paragraphData struct {
    XMLName    xml.Name             `xml:"w:p"`
    Properties *paragraphProperties `xml:"w:pPr,omitempty"`
    Runs       []paragraphRun       `xml:"w:r,omitempty"`
}

type documentBodyData struct {
    XMLName    xml.Name        `xml:"w:body"`
    Paragraphs []paragraphData `xml:"w:p"`
    SectPr     *sectPr         `xml:"w:sectPr"`
}


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


type xmlRootDocument struct {
    XMLName xml.Name         `xml:"w:document"`
    XmlnsWp string           `xml:"xmlns:wp,attr"`
    XmlnsA  string           `xml:"xmlns:a,attr"`
    XmlnsPic string          `xml:"xmlns:pic,attr"`
    XmlnsR  string           `xml:"xmlns:r,attr"`
    XmlnsW  string           `xml:"xmlns:w,attr"`

    Body documentBodyData `xml:"w:body"`
}


type Document interface {
    AddText(style string, textData string, formatOptions ...string)
    AddNewLine()
    AddImage(filepath string) error
    renderContent(w io.Writer) error
    getImages() map[string][]byte
    getImageContentTypes() map[string]string
    getImageRelationships() []relationship
}


type DocxDocument struct {
    content           []paragraphData
    images            map[string][]byte
    imageContentTypes map[string]string
    imageRels         []relationship
    imageCounter      uint
    lastRID           int
}


func NewDocxDocument() *DocxDocument {
    return &DocxDocument{
        content:           []paragraphData{},
        images:            make(map[string][]byte),
        imageContentTypes: make(map[string]string),
        imageRels:         []relationship{},
        imageCounter:      0,
        lastRID:           1,
    }
}


func (d *DocxDocument) nextRID() string {
    d.lastRID++
    return fmt.Sprintf("rId%d", d.lastRID)
}


func (d *DocxDocument) AddText(style string, textData string, formatOptions ...string) {
    runProps := runProperties{}
    runText := paragraphRunText{Text: textData, Space: "preserve"}
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


    canAppend := false
    if len(d.content) > 0 {
        lastParaIndex := len(d.content) - 1
        lastPara := &d.content[lastParaIndex]




        hasDrawing := false
        for _, r := range lastPara.Runs {
            if r.Drawing != nil {
                hasDrawing = true
                break
            }
        }

        if !hasDrawing && (len(lastPara.Runs) > 0 || lastPara.Properties != nil) {
            lastStyle := StyleNormal
            if lastPara.Properties != nil && lastPara.Properties.Style != nil {
                lastStyle = lastPara.Properties.Style.Val
            }

            currentStyle := style
            if currentStyle == "" {
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

        paraProps := paragraphProperties{}
        var finalParaProps *paragraphProperties


        validStyle := style
        styleIsSet := false
        switch style {
        case StyleHeading1, StyleHeading2, StyleHeading3, StyleHeading4:
            paraProps.Style = &paragraphStyle{Val: style}
            finalParaProps = &paraProps
            styleIsSet = true
        case StyleNormal, "":
            validStyle = StyleNormal


            if style == StyleNormal {
                paraProps.Style = &paragraphStyle{Val: StyleNormal}
                finalParaProps = &paraProps
                styleIsSet = true
            }
        default:

            if strings.TrimSpace(style) != "" {
                paraProps.Style = &paragraphStyle{Val: style}
                finalParaProps = &paraProps
                styleIsSet = true
            }
        }


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


func (d *DocxDocument) AddNewLine() {


    d.content = append(d.content, paragraphData{})
}


func (d *DocxDocument) AddImage(filePath string) error {
    d.imageCounter++
    imgID := d.imageCounter
    uniquePicID := imgID
    rID := d.nextRID()

    imgBytes, err := os.ReadFile(filePath)
    if err != nil {
        return fmt.Errorf("failed to read image file %s: %w", filePath, err)
    }


    imgDataReader := bytes.NewReader(imgBytes)

    imgConfig, format, err := image.DecodeConfig(imgDataReader)
    if err != nil {

        _, _ = imgDataReader.Seek(0, io.SeekStart)
        _, format, err = image.Decode(imgDataReader)
        if err != nil {
            return fmt.Errorf("failed to decode image config or data for %s: %w", filePath, err)
        }

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
        imgExt = ".jpg"
    case "png":
        contentType = "image/png"
        imgExt = ".png"
    case "gif":
        contentType = "image/gif"
        imgExt = ".gif"
    default:
        return fmt.Errorf("unsupported image format: %s for file %s", format, filePath)
    }


    imgFileName := fmt.Sprintf("image%d%s", imgID, imgExt)


    d.images[imgFileName] = imgBytes
    d.imageContentTypes[contentType] = imgExt[1:]


    d.imageRels = append(d.imageRels, relationship{
        ID:     rID,
        Type:   "http:
        Target: fmt.Sprintf("media/%s", imgFileName),
    })


    widthEMU := int64(imgConfig.Width) * emusPerPixel
    heightEMU := int64(imgConfig.Height) * emusPerPixel


    descr := "Inserted Picture"
    drawing := Drawing{
        Inline: Inline{
            DistT: 0, DistB: 0, DistL: 0, DistR: 0,
            Extent:       extent{Cx: widthEMU, Cy: heightEMU},
            EffectExtent: effectExtent{L: 0, T: 0, R: 0, B: 0},
            DocPr: DocProperties{
                ID:    imgID,
                Name:  fmt.Sprintf("Picture %d", imgID),
                Descr: descr,
            },
            CNvGraphicFramePr: CnvGraphicFrameProperties{
                GraphicFrame: graphicFrameLocks{
                    NoChangeAspect: 1,
                },
            },
            Graphic: graphic{
                GraphicData: graphicData{
                    URI: "http:
                    Pic: pic{
                        NvPicPr: nonVisualPicProperties{
                            CNvPr: struct {
                                XMLName xml.Name `xml:"pic:cNvPr"`
                                ID      uint     `xml:"id,attr"`
                                Name    string   `xml:"name,attr"`
                                Descr   string   `xml:"descr,attr,omitempty"`
                            }{ID: uniquePicID, Name: imgFileName, Descr: descr},
                            CNvPicPr: cNvPicPr{
                                PicLocks: picLocks{
                                    NoChangeAspect:     1,
                                    NoChangeArrowheads: 1,
                                },
                            },
                        },
                        BlipFill: blipFill{
                            Blip: blip{
                                Embed:  rID,
                                Cstate: "print",
                            },
                            SrcRect: &srcRect{},
                            Stretch: stretch{
                                FillRectangle: &struct {
                                    XMLName xml.Name `xml:"a:fillRect"`
                                }{},
                            },
                        },
                        SpPr: shapeProperties{
                            BwMode: "auto",
                            Xfrm: transform2D{
                                Offset:  point2D{X: 0, Y: 0},
                                Extents: extents{Cx: widthEMU, Cy: heightEMU},
                            },
                            PrstGeom: presetGeometry{
                                Prst: "rect",
                                AVList: &struct {
                                    XMLName xml.Name `xml:"a:avLst"`
                                }{},
                            },
                            NoFill: &noFill{},
                            Ln: &ln{
                                NoFill: &noFill{},




                            },
                        },
                    },
                },
            },
        },
    }


    imgRun := paragraphRun{Drawing: &drawing}

    para := paragraphData{Runs: []paragraphRun{imgRun}}
    d.content = append(d.content, para)

    return nil
}


func (d *DocxDocument) renderContent(w io.Writer) error {

    doc := xmlRootDocument{
        XmlnsWp:  "http:
        XmlnsA:   "http:
        XmlnsPic: "http:
        XmlnsR:   "http:
        XmlnsW:   "http:
        Body: documentBodyData{
            Paragraphs: d.content,

            SectPr: &sectPr{
                PgSz: pgSz{W: 12240, H: 15840},
                PgMar: pgMar{Top: 1440, Right: 1440, Bottom: 1440, Left: 1440, Header: 720, Footer: 720, Gutter: 0},
                Cols: struct {
                    XMLName xml.Name `xml:"w:cols"`
                    Space   uint     `xml:"w:space,attr"`
                }{Space: 720},
                DocGrid: struct {
                    XMLName   xml.Name `xml:"w:docGrid"`
                    LinePitch uint     `xml:"w:linePitch,attr"`
                }{LinePitch: 360},
            },
        },
    }


    _, err := w.Write([]byte(xml.Header))
    if err != nil {
        return fmt.Errorf("failed to write xml header: %w", err)
    }


    encoder := xml.NewEncoder(w)
    encoder.Indent("", "  ")
    err = encoder.Encode(doc)
    if err != nil {
        return fmt.Errorf("failed to encode document content: %w", err)
    }

    return nil
}


func (d *DocxDocument) getImages() map[string][]byte {
    return d.images
}

func (d *DocxDocument) getImageContentTypes() map[string]string {
    return d.imageContentTypes
}

func (d *DocxDocument) getImageRelationships() []relationship {
    return d.imageRels
}

