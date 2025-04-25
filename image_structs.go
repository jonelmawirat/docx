package docx

import "encoding/xml"

// --- DocProperties ---
type DocProperties struct {
    XMLName xml.Name `xml:"wp:docPr"`
    ID      uint     `xml:"id,attr"`
    Name    string   `xml:"name,attr"`
    Descr   string   `xml:"descr,attr,omitempty"` // Added Descr
}

// --- graphicFrameLocks ---
type graphicFrameLocks struct {
    XMLName        xml.Name `xml:"a:graphicFrameLocks"`
    XmlnsA         string   `xml:"xmlns:a,attr,omitempty"` // Keep xmlns if needed, but often inherited
    NoChangeAspect uint     `xml:"noChangeAspect,attr"`
}

// --- CnvGraphicFrameProperties ---
type CnvGraphicFrameProperties struct {
    XMLName      xml.Name          `xml:"wp:cNvGraphicFramePr"`
    GraphicFrame graphicFrameLocks `xml:"a:graphicFrameLocks"`
}

// --- point2D ---
type point2D struct {
    XMLName xml.Name `xml:"a:off"`
    X       int64    `xml:"x,attr"`
    Y       int64    `xml:"y,attr"`
}

// --- extents ---
type extents struct {
    XMLName xml.Name `xml:"a:ext"`
    Cx      int64    `xml:"cx,attr"`
    Cy      int64    `xml:"cy,attr"`
}

// --- transform2D ---
type transform2D struct {
    XMLName xml.Name `xml:"a:xfrm"`
    Offset  point2D  `xml:"a:off"`
    Extents extents  `xml:"a:ext"`
}

// --- Preset Geometry ---
type presetGeometry struct {
    XMLName xml.Name `xml:"a:prstGeom"`
    Prst    string   `xml:"prst,attr"`
    AVList  *struct { // Use pointer to empty struct for <a:avLst/>
        XMLName xml.Name `xml:"a:avLst"`
    } `xml:"a:avLst"`
}

// --- blip ---
type blip struct {
    XMLName xml.Name `xml:"a:blip"`
    // Removed XmlnsR string   `xml:"xmlns:r,attr"` // Rely on top-level xmlns definition
    Embed  string `xml:"r:embed,attr"`
    Cstate string `xml:"cstate,attr,omitempty"`
}

// --- srcRect (empty element) ---
type srcRect struct {
    XMLName xml.Name `xml:"a:srcRect"`
}

// --- stretch ---
type stretch struct {
    XMLName       xml.Name `xml:"a:stretch"`
    FillRectangle *struct { // Use pointer to empty struct for self-closing tag <a:fillRect/>
        XMLName xml.Name `xml:"a:fillRect"`
    } `xml:"a:fillRect"`
}

// --- blipFill (Corrected) ---
type blipFill struct {
    XMLName xml.Name `xml:"pic:blipFill"`
    Blip    blip     `xml:"a:blip"`
    SrcRect *srcRect `xml:"a:srcRect,omitempty"` // Added srcRect (pointer for empty tag)
    Stretch stretch  `xml:"a:stretch"`
}

// --- Structs for ShapeProperties' Line Properties ---
type noFill struct { // Represents <a:noFill/>
    XMLName xml.Name `xml:"a:noFill"`
}

type headEnd struct { // Represents <a:headEnd/>
    XMLName xml.Name `xml:"a:headEnd"`
    // Add attributes like type, w, len if needed
}

type tailEnd struct { // Represents <a:tailEnd/>
    XMLName xml.Name `xml:"a:tailEnd"`
    // Add attributes like type, w, len if needed
}

type ln struct { // Represents <a:ln>
    XMLName xml.Name `xml:"a:ln"`
    NoFill  *noFill  `xml:"a:noFill,omitempty"` // Optional NoFill within ln
    Miter   *struct {
        XMLName xml.Name `xml:"a:miter"`
        Lim     string   `xml:"lim,attr"`
    } `xml:"a:miter,omitempty"`
    HeadEnd *headEnd `xml:"a:headEnd,omitempty"`
    TailEnd *tailEnd `xml:"a:tailEnd,omitempty"`
}

// --- shapeProperties (Corrected) ---
type shapeProperties struct {
    XMLName  xml.Name       `xml:"pic:spPr"`
    BwMode   string         `xml:"bwMode,attr,omitempty"` // Added bwMode
    Xfrm     transform2D    `xml:"a:xfrm"`
    PrstGeom presetGeometry `xml:"a:prstGeom"`
    NoFill   *noFill        `xml:"a:noFill,omitempty"` // Added NoFill (pointer for empty tag)
    Ln       *ln            `xml:"a:ln,omitempty"`     // Added Ln (Line properties)
}

// --- picLocks (Added) ---
type picLocks struct {
    XMLName            xml.Name `xml:"a:picLocks"`
    NoChangeAspect     uint     `xml:"noChangeAspect,attr,omitempty"`
    NoChangeArrowheads uint     `xml:"noChangeArrowheads,attr,omitempty"`
}

// --- cNvPicPr (Corrected) ---
type cNvPicPr struct {
    XMLName  xml.Name `xml:"pic:cNvPicPr"`
    PicLocks picLocks `xml:"a:picLocks"` // Use the picLocks struct
}

// --- nonVisualPicProperties (Corrected) ---
type nonVisualPicProperties struct {
    XMLName xml.Name `xml:"pic:nvPicPr"`
    CNvPr   struct { // Keep this embedded struct for cNvPr
        XMLName xml.Name `xml:"pic:cNvPr"`
        ID      uint     `xml:"id,attr"`
        Name    string   `xml:"name,attr"`
        Descr   string   `xml:"descr,attr,omitempty"` // Added Descr
    } `xml:"pic:cNvPr"`
    CNvPicPr cNvPicPr `xml:"pic:cNvPicPr"` // Use the modified cNvPicPr struct
}

// --- pic (Corrected) ---
type pic struct {
    XMLName  xml.Name               `xml:"pic:pic"`
    XmlnsPic string                 `xml:"xmlns:pic,attr,omitempty"` // Namespace can be inherited
    NvPicPr  nonVisualPicProperties `xml:"pic:nvPicPr"`              // Uses updated nonVisualPicProperties
    BlipFill blipFill               `xml:"pic:blipFill"`             // Uses updated blipFill
    SpPr     shapeProperties        `xml:"pic:spPr"`                 // Uses updated shapeProperties
}

// --- graphicData (Corrected) ---
type graphicData struct {
    XMLName xml.Name `xml:"a:graphicData"`
    URI     string   `xml:"uri,attr"`
    Pic     pic      `xml:"pic:pic"` // Uses updated pic
}

// --- graphic (Corrected) ---
type graphic struct {
    XMLName     xml.Name    `xml:"a:graphic"`
    XmlnsA      string      `xml:"xmlns:a,attr,omitempty"` // Namespace can be inherited
    GraphicData graphicData `xml:"a:graphicData"`          // Uses updated graphicData
}

// --- extent ---
type extent struct {
    XMLName xml.Name `xml:"wp:extent"`
    Cx      int64    `xml:"cx,attr"`
    Cy      int64    `xml:"cy,attr"`
}

// --- effectExtent ---
type effectExtent struct {
    XMLName xml.Name `xml:"wp:effectExtent"`
    L       int64    `xml:"l,attr"`
    T       int64    `xml:"t,attr"`
    R       int64    `xml:"r,attr"`
    B       int64    `xml:"b,attr"`
}

// --- Inline (Corrected) ---
type Inline struct {
    XMLName           xml.Name                `xml:"wp:inline"`
    DistT             uint                    `xml:"distT,attr"`
    DistB             uint                    `xml:"distB,attr"`
    DistL             uint                    `xml:"distL,attr"`
    DistR             uint                    `xml:"distR,attr"`
    DocPr             DocProperties           `xml:"wp:docPr"` // Uses updated DocProperties
    CNvGraphicFramePr CnvGraphicFrameProperties `xml:"wp:cNvGraphicFramePr"` // Uses updated CnvGraphicFrameProperties
    Extent            extent                  `xml:"wp:extent"`
    EffectExtent      effectExtent            `xml:"wp:effectExtent"`
    Graphic           graphic                 `xml:"a:graphic"` // Uses updated graphic
}

// --- Drawing (Corrected) ---
type Drawing struct {
    XMLName xml.Name `xml:"w:drawing"`
    Inline  Inline   `xml:"wp:inline"` // Uses updated Inline
}
